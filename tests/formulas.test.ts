import { describe, it, expect } from "vitest";
import { WorkbookModel } from "../src/core/model";
import { keyOf, parseRange } from "../src/core/address";
import {
  completeFunction,
  formulaContext,
} from "../src/editor/formula-context";

describe("Formula editing context", () => {
  it("tracks nested arguments without treating quoted or array commas as separators", () => {
    const value = '=IF(A1, SUM({1,2;3,4}, "a,""b", \'Sales, Q1\'!B2), ';
    expect(formulaContext(value, value.length).call).toEqual({
      name: "IF",
      argument: 2,
    });
    const nested = '=SUM((1+2), IF(A1,"yes,no",';
    expect(formulaContext(nested, nested.length).call).toEqual({
      name: "IF",
      argument: 2,
    });
    expect(formulaContext('=IF("SU', 7).token).toBeUndefined();
    expect(formulaContext("='Sales SU", 10).token).toBeUndefined();
    expect(formulaContext("=SUM!A1", 4).token).toBeUndefined();
    expect(formulaContext("=Sheet1!SU", 10).token).toBeUndefined();
    expect(formulaContext("'=SU", 4)).toEqual({});
  });
  it("completes at the caret without overwriting suffixes or duplicating parentheses", () => {
    const value = "=IF(A1,SUM(B1:B3),0)";
    const caret = value.indexOf("SUM") + 2;
    const context = formulaContext(value, caret);
    expect(context.token?.text).toBe("SU");
    expect(completeFunction(value, context.token!, "SUMIF", false)).toEqual({
      value: "=IF(A1,SUMIF(B1:B3),0)",
      caret: 13,
    });
    expect(
      completeFunction("=TO", formulaContext("=TO", 3).token!, "TODAY", true),
    ).toEqual({
      value: "=TODAY()",
      caret: 8,
    });
  });
});
const cases: [string, unknown][] = [
  ['MATCH(">5",{"abc";">5"},0)', 2],
  ['MATCH("123",{123;"123"},0)', 2],
  ['MATCH("=a*",{"abc";"=ABC"},0)', 2],
  ['MATCH("~*",{"abc";"*"},0)', 2],
  ['MATCH("",A99:A100,0)', 1],
  ['VLOOKUP(">=x",{"z",1;">=x",2},2,FALSE)', 2],
  ['HLOOKUP("<>x",{"abc","<>x";1,2},2,FALSE)', 2],
  ['XLOOKUP(">*",{"abc";">x"},{1;2},,2)', 2],
  ['XMATCH("=a?",{"abc";"=Ab"},2)', 2],
  ["SUM(1,2,3)", 6],
  ["SUMSQ(2,3)", 13],
  ["PRODUCT(2,3,4)", 24],
  ["AVERAGE(2,4,6)", 4],
  ["MEDIAN(1,9,2,8)", 5],
  ['COUNT(1,"text",TRUE)', 2],
  ['COUNTA(1,"",TRUE)', 3],
  ["MIN(3,1,9)", 1],
  ["MAX(3,1,9)", 9],
  ["ABS(-2)", 2],
  ["ROUND(-1.25,1)", -1.3],
  ["ROUNDUP(1.231,2)", 1.24],
  ["ROUNDDOWN(-1.239,2)", -1.23],
  ["TRUNC(-2.8)", -2],
  ["MOD(-3,2)", 1],
  ["QUOTIENT(9,2)", 4],
  ["CEILING.MATH(-4.3)", -4],
  ["FLOOR.MATH(-4.3)", -5],
  ["MROUND(10,3)", 9],
  ["POWER(3,3)", 27],
  ["FACT(6)", 720],
  ["COMBIN(10,3)", 120],
  ["GCD(12,18)", 6],
  ["LCM(6,8)", 24],
  ["EVEN(-3)", -4],
  ["ODD(0)", 1],
  ["SQRT(81)", 9],
  ["LOG(100)", 2],
  ["SIGN(-3)", -1],
  ["INT(-2.1)", -3],
  ["VAR.S(1,2,3)", 1],
  ["STDEV.P(1,1,1)", 0],
  ["LARGE({1,5,3},2)", 3],
  ["SMALL({1,5,3},2)", 3],
  ["PERCENTILE.INC({1,2,3,4},0.5)", 2.5],
  ["RANK.EQ(5,{7,5,5,2})", 2],
  ["MODE.SNGL(1,2,2,3)", 2],
  ["SUMPRODUCT({1,2},{3,4})", 11],
  ["IF(FALSE,1/0,7)", 7],
  ['IFERROR(SQRT(-1),"fallback")', "fallback"],
  ["IFNA(NA(),4)", 4],
  ["IFERROR(SUM({1,#N/A}),42)", 42],
  ["IFS(FALSE,1,TRUE,2)", 2],
  ['SWITCH("a","b",1,"a",2,3)', 2],
  ['CHOOSE(2,"x","y")', "y"],
  ["AND(TRUE,1)", true],
  ["OR(FALSE,0)", false],
  ["XOR(TRUE,TRUE,TRUE)", true],
  ["NOT(0)", true],
  ["ISBLANK(A99)", true],
  ["ISNUMBER(3)", true],
  ['ISTEXT("3")', true],
  ["ISERR(NA())", false],
  ["ISERROR(1/0)", true],
  ["ISNA(NA())", true],
  ["ISEVEN(-4)", true],
  ["ISODD(-3)", true],
  ["TYPE(TRUE)", 4],
  ['N("hello")', 0],
  ["T(3)", ""],
  ['LEFT("OnlineExcel",6)', "Online"],
  ['RIGHT("abc",0)', ""],
  ['MID("abcdef",2,3)', "bcd"],
  ['LEN("hello")', 5],
  ['UPPER("hello")', "HELLO"],
  ['LOWER("ABC")', "abc"],
  ['TRIM(" a   b ")', "a b"],
  ['PROPER("hello WORLD")', "Hello World"],
  ['FIND("B","aBc")', 2],
  ['SEARCH("b?","aBC")', 2],
  ['REPLACE("abcdef",2,3,"X")', "aXef"],
  ['SUBSTITUTE("a-a-a","a","X",2)', "a-X-a"],
  ['REPT("ab",3)', "ababab"],
  ['CONCAT("a","b",3)', "ab3"],
  ['TEXTJOIN("-",TRUE,"a","","b")', "a-b"],
  ['EXACT("A","a")', false],
  ["CHAR(65)", "A"],
  ['UNICODE("中")', 20013],
  ["UNICHAR(20013)", "中"],
  ['VALUE("1,234.5%")', 12.345],
  ['NUMBERVALUE("1.234,56",",",".")', 1234.56],
  ['TEXT(0.25,"0%")', "25%"],
  ["FIXED(1234.567,2,FALSE)", "1,234.57"],
  ["DATE(2024,1,1)", 45292],
  ["YEAR(DATE(2024,1,1))", 2024],
  ["MONTH(DATE(2024,2,1))", 2],
  ["DAY(60)", 29],
  ["TIME(12,0,0)", 0.5],
  ["HOUR(0.5)", 12],
  ["MINUTE(TIME(12,30,0))", 30],
  ["DAYS(DATE(2024,1,3),DATE(2024,1,1))", 2],
  ["WEEKDAY(DATE(2024,1,1),2)", 1],
  ["ISOWEEKNUM(DATE(2024,1,1))", 1],
  ["EOMONTH(DATE(2024,1,31),1)", 45351],
  ["EDATE(DATE(2024,1,31),1)", 45351],
  ['DATEVALUE("2024-01-01")', 45292],
  ['TIMEVALUE("12:30 PM")', 12.5 / 24],
  ["NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,7))", 5],
  ["WORKDAY(DATE(2024,1,5),1)", 45299],
  ["INDEX({10,20;30,40},2,2)", 40],
  ['MATCH("b",{"a","b","c"},0)', 2],
  ["XMATCH(15,{10,20,30},-1)", 1],
  ['XLOOKUP(20,{10;20;30},{"a";"b";"c"})', "b"],
  ["VLOOKUP(2,{1,10;2,20;3,30},2,FALSE)", 20],
  ["HLOOKUP(2,{1,2,3;10,20,30},2,FALSE)", 20],
  ["ROWS({1,2;3,4})", 2],
  ["COLUMNS({1,2;3,4})", 2],
  ["ADDRESS(2,3,1)", "$C$2"],
  ["SUM(SEQUENCE(3,2))", 21],
  ["SUM(FILTER({1;2;3},{TRUE;FALSE;TRUE}))", 4],
  ["SUM(UNIQUE({1;1;2}))", 3],
  ["SUM(TAKE({1;2;3},2))", 3],
  ["SUM(DROP({1;2;3},1))", 5],
  ["SUM(CHOOSECOLS({1,2,3;4,5,6},1,3))", 14],
  ["SUM(VSTACK({1,2},{3,4}))", 10],
  ["PMT(0,10,1000)", -100],
  ["FV(0,10,-100)", 1000],
  ["PV(0,10,-100)", 1000],
  ["NPER(0,-100,1000)", 10],
  ["NPV(0.1,110,121)", 200],
  ["IRR({-100;110})", 0.1],
  ["SLN(1000,100,3)", 300],
  ["SYD(1000,100,3,1)", 450],
  ["EFFECT(0.12,12)", 1.01 ** 12 - 1],
  ["NOMINAL(1.01^12-1,12)", 0.12],
];
describe("formula reference cases", () => {
  it.each([
    ["=IFERROR({#N/A;#N/A},{10;20})", [[10], [20]]],
    ["=IFNA({#N/A;1;#DIV/0!},{10;20;30})", [[10], [1], [{ error: "#DIV/0!" }]]],
    ["=IFERROR(1/0,{10;20})", [[10], [20]]],
    [
      "=IFERROR({#N/A;#N/A},{10,20})",
      [
        [10, 20],
        [10, 20],
      ],
    ],
    [
      "=IF({TRUE;FALSE},{1,2},{3,4})",
      [
        [1, 2],
        [3, 4],
      ],
    ],
    ["=IF({TRUE;TRUE;TRUE},{1;2},0)", [[1], [2], [{ error: "#N/A" }]]],
    ["=IF({TRUE;TRUE},{1;2},SEQUENCE(100))", [[1], [2]]],
    ["=IFERROR({1;2},SEQUENCE(100))", [[1], [2]]],
    ["=IF({TRUE;FALSE},{1;#N/A},{#DIV/0!;2})", [[1], [2]]],
  ] as [string, unknown[][]][])(
    "selects conditional array elements and dimensions: %s",
    (formula, expected) => {
      const m = new WorkbookModel({
        sheets: [{ name: "Data", rows: 20, columns: 10 }],
      });
      const id = m.sheets[0].meta.id;
      m.execute([
        { type: "setFormula", sheetId: id, row: 0, column: 3, formula },
      ]);
      const actual = expected.map((row, r) =>
        row.map((_, c) => m.engine.get(id, keyOf(r, c + 3))),
      );
      expect(actual).toEqual(expected);
      expect(m.getDiagnostics()).toEqual([]);
    },
  );
  for (const [formula, expected] of cases)
    it(formula, () => {
      const model = new WorkbookModel({
          sheets: [{ name: "Test", rows: 200, columns: 30 }],
        }),
        id = model.sheets[0].meta.id;
      model.execute([
        {
          type: "setFormula",
          sheetId: id,
          row: 0,
          column: 0,
          formula: "=" + formula,
        },
      ]);
      const actual = model.engine.get(id, 0);
      if (typeof expected === "number") expect(actual).toBeCloseTo(expected, 8);
      else expect(actual).toEqual(expected);
    });
  it("evaluates conditional aggregates with wildcard and multiple criteria", () => {
    const m = new WorkbookModel({
        sheets: [{ name: "Data", rows: 100, columns: 20 }],
      }),
      id = m.sheets[0].meta.id;
    m.execute([
      {
        type: "setValues",
        sheetId: id,
        range: parseRange("A1:C4"),
        values: [
          ["East", 10, 1],
          ["East", 20, 2],
          ["West", 30, 3],
          ["East", 40, 4],
        ],
      },
      {
        type: "setFormula",
        sheetId: id,
        row: 0,
        column: 4,
        formula: '=SUMIFS(B1:B4,A1:A4,"E*",C1:C4,">1")',
      },
    ]);
    expect(m.engine.get(id, keyOf(0, 4))).toBe(60);
  });
  it("broadcasts arrays and follows dynamic spill output", () => {
    const m = new WorkbookModel({
        sheets: [{ name: "Data", rows: 100, columns: 20 }],
      }),
      id = m.sheets[0].meta.id;
    m.execute([
      {
        type: "setFormula",
        sheetId: id,
        row: 0,
        column: 0,
        formula: "=SEQUENCE(3)*10",
      },
      {
        type: "setFormula",
        sheetId: id,
        row: 0,
        column: 2,
        formula: "=FILTER(A1#,A1#>10)",
      },
    ]);
    expect(m.engine.get(id, keyOf(1, 2))).toBe(30);
  });
});
import {
  referenceInsertion,
  cycleReference,
} from "../src/editor/formula-reference";

it("inserts and cycles formula references without replacing quoted text or function names", () => {
  expect(referenceInsertion("=SUM(", 5, 5)).toEqual({ start: 5, end: 5 });
  expect(referenceInsertion('="A1"', 3, 3)).toBeUndefined();
  expect(referenceInsertion("=SUM(12)", 6, 6)).toBeUndefined();
  expect(cycleReference("=LOG10(2)", 5)).toBeUndefined();
  const values = [
    "='O''Brien'!A1:B2",
    "='O''Brien'!$A$1:$B$2",
    "='O''Brien'!A$1:B$2",
    "='O''Brien'!$A1:$B2",
    "='O''Brien'!A1:B2",
  ];
  for (let i = 0; i < values.length - 1; i++)
    expect(cycleReference(values[i], values[i].length)?.value).toBe(
      values[i + 1],
    );
});
