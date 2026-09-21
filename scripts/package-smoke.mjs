import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { resolve, join } from "node:path";
import { execFileSync } from "node:child_process";
const { version } = JSON.parse(readFileSync("package.json", "utf8"));
const requested = process.argv[2] ?? `artifacts/onlineexcel-${version}.tgz`;
const packageSpec = requested.startsWith("onlineexcel@")
  ? requested
  : resolve(requested);
const directory = mkdtempSync(join(tmpdir(), "onlineexcel-package-"));
execFileSync(
  "npm",
  [
    "install",
    "--prefix",
    directory,
    "--ignore-scripts",
    "--no-audit",
    "--no-fund",
    "--registry=https://registry.npmjs.org/",
    packageSpec,
  ],
  { stdio: "inherit" },
);
writeFileSync(
  join(directory, "smoke.mjs"),
  `
import * as publicAPI from 'onlineexcel';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';
const { listFunctions, createWorkbook, mountEditor, parseRange } = publicAPI;
const context = { document: { currentScript: { src: 'https://example.test/onlineexcel.js' } }, URL };
runInNewContext(readFileSync(new URL('./node_modules/onlineexcel/dist/onlineexcel.js', import.meta.url), 'utf8'), context);
if (Object.keys(publicAPI).sort().join() !== Object.keys(context.OnlineExcel).sort().join()) throw new Error('ESM/IIFE export mismatch');
console.log('PASS: ESM/IIFE runtime export parity');
import { WorkbookModel } from 'onlineexcel/core';
import { readXlsx, writeXlsx } from 'onlineexcel/xlsx';
import { mountEditor as editorEntry } from 'onlineexcel/editor';
if (typeof createWorkbook !== 'function' || typeof mountEditor !== 'function' || typeof editorEntry !== 'function') throw new Error('Missing exports');
if (listFunctions().length !== 187) throw new Error('Function manifest differs');
const model = new WorkbookModel(); const id = model.sheets[0].meta.id;
model.execute([{ type:'setValues', sheetId:id, range:parseRange('A1:C1'), values:[[2,3,'=SUM(A1:B1)']] }]);
const result = await readXlsx((await writeXlsx(model.snapshot())).data);
const next = new WorkbookModel({snapshot:result.snapshot});
if (next.region(next.sheets[0].meta.id,parseRange('C1')).cells[0].value !== 5) throw new Error('Packaged IO failed');
console.log('PASS: installed ESM, core, editor and XLSX exports');
`,
);
execFileSync("node", [join(directory, "smoke.mjs")], { stdio: "inherit" });
writeFileSync(
  join(directory, "consumer.ts"),
  `
import { createWorkbook, mountEditor, createPersistence, createIndexedDBStorage, supportedLocales, type Locale, type WorkbookSnapshot } from 'onlineexcel';
import { type Locale as EditorLocale } from 'onlineexcel/editor';
import { readXlsx } from 'onlineexcel/xlsx';
async function embed(container: HTMLElement) {
 const workbook = await createWorkbook({workerUrl:'/assets/worker.js'});
 const locale: Locale = 'ja-JP';
 const editorLocale: EditorLocale = 'ko-KR';
 const editor = mountEditor(container,{workbook,locale});
 const persistence = createPersistence(workbook,{key:'test',storage:createIndexedDBStorage('host')});
 await persistence.ready;
 await editor.setOptions({persistence,zoom:1.25});
 await editor.setZoom(1.5);
 const zoom: number = editor.getZoom();
 for (const language of supportedLocales) mountEditor(container,{workbook,locale:language}).destroy();
 mountEditor(container,{workbook,locale:editorLocale}).destroy();
 const dirty: boolean = editor.getEditState().dirty;
 await editor.commitEdit();
 const id = (await workbook.getMetadata()).sheets[0].id;
 await workbook.getNavigationTarget(id,'A1','down');
 await workbook.setValidation(id,'A1:A10',{type:'list',values:['Yes','No']});
 await workbook.setConditionalFormat(id,'B1:B10',{operator:'gt',value:5,style:{background:'#ff0000'}});
 await workbook.setStyle(id,'A1:A10',{locked:false});
 await workbook.protectSheet(id);
 await workbook.protectSheet(id,false);
 await workbook.importCsv('Code,Rate\\n001,25%',{sheetId:id,columns:['text','percent'],header:true});
 const copy = await workbook.duplicateSheet(id,'Copy');
 await workbook.setSheetHidden(copy,true);
 await workbook.defineName('Input',id,'A1'); await workbook.renameName('Input','Input2'); await workbook.deleteName('Input2');
 await editor.setOptions({locale:'zh-TW',readOnly:true,toolbarItems:['saveXlsx']});
 const checkpoint = await workbook.createSavePoint();
 await workbook.withOptions({origin:'host',operationId:'save-1',timeout:3000}).getMetadata();
 await workbook.setRecords(id,[{code:'001'}],[{key:'code'}]);
 for await (const records of workbook.readRecords(id,'A1',[{key:'code'}])) console.log(records);
 for await (const chunk of workbook.readChunks(id,'A1')) console.log(chunk.revision);
 await workbook.writeChunks(id,'A1',[[[1]]]);
 const snapshot: WorkbookSnapshot = checkpoint.snapshot;
 workbook.on('change', event => console.log(event.revision));
 await workbook.importJSON(snapshot);
 const exported = await workbook.exportXlsx();
 await readXlsx(exported.data);
 await editor.destroy({commit:true}); persistence.dispose(); await workbook.dispose();
}
`,
);
execFileSync(
  resolve("node_modules/.bin/tsc"),
  [
    join(directory, "consumer.ts"),
    "--strict",
    "--noEmit",
    "--skipLibCheck",
    "--moduleResolution",
    "bundler",
    "--module",
    "esnext",
    "--target",
    "es2022",
    "--lib",
    "es2022,dom,dom.iterable,webworker",
  ],
  { stdio: "inherit" },
);
console.log("PASS: installed TypeScript declarations");
