// Trusted application code. The Worker loads this module before opening the workbook.
export const functions = [
  {
    name: "ACME.DOUBLE",
    minArgs: 1,
    maxArgs: 1,
    signature: "value",
    description: "Double a number and spill the original value beside it.",
    evaluate: ([value]) => [[Number(value) * 2, Number(value)]],
  },
];
export function validate({ sheets, getRegion }) {
  const sheet = sheets[0];
  const cell = getRegion(sheet.id, "A1").cells[0];
  if (typeof cell?.value === "number" && cell.value < 0)
    return [
      {
        code: "NEGATIVE_AMOUNT",
        severity: "error",
        sheetId: sheet.id,
        range: "A1",
        message: "Amount must be nonnegative",
      },
    ];
}
