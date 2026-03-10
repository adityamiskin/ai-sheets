export const DEFAULT_WORKBOOK_NAME = "Book1.xlsx";

export const DEFAULT_SNAPSHOT = JSON.stringify({
  id: "workbook-1",
  locale: "enUS",
  name: "Book1",
  sheetOrder: ["sheet-1"],
  sheets: {
    "sheet-1": {
      id: "sheet-1",
      name: "Sheet1",
      rowCount: 100,
      columnCount: 26,
      cellData: {},
      defaultColumnWidth: 88,
      defaultRowHeight: 24,
      showGridlines: 1,
    },
  },
});
