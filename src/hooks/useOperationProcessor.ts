import { useQuery, useMutation } from "convex/react";
import { api } from "../../convex/_generated/api";
import { useEffect, useRef } from "react";
import type { Id } from "../../convex/_generated/dataModel";
import * as GC from "@mescius/spread-sheets";
import "@mescius/spread-sheets-charts";
import type { SpreadsheetApi } from "@/components/spreadsheet/SpreadsheetEditor";

/**
 * Subscribes to pending operations and applies them to the SpreadJS workbook.
 */
const MIN_SHEET_COLUMN_COUNT = 200;
const MIN_SHEET_ROW_COUNT = 1000;

export function useOperationProcessor(
  workbookId: Id<"workbooks"> | null,
  spread: SpreadsheetApi | null
) {
  const pendingOps = useQuery(
    api.operations.getPending,
    workbookId ? { workbookId } : "skip"
  );
  const markApplied = useMutation(api.operations.markApplied);
  const markFailed = useMutation(api.operations.markFailed);
  const processingRef = useRef(new Set<string>());

  useEffect(() => {
    if (!pendingOps || !spread || pendingOps.length === 0) return;

    for (const op of pendingOps) {
      const opId = op._id as string;
      if (processingRef.current.has(opId)) continue;
      processingRef.current.add(opId);

      (async () => {
        try {
          const params = JSON.parse(op.params);
          applyOperation(spread, op.type, params);
          await markApplied({ id: op._id });
        } catch (e) {
          const errMsg = e instanceof Error ? e.message : "Unknown error";
          console.error(`Failed to apply operation ${op.type}:`, errMsg);
          await markFailed({ id: op._id, error: errMsg });
        } finally {
          processingRef.current.delete(opId);
        }
      })();
    }
  }, [pendingOps, spread, markApplied, markFailed]);
}

/**
 * Applies a single operation to the SpreadJS workbook.
 */
function applyOperation(
  spread: SpreadsheetApi,
  type: string,
  params: Record<string, unknown>
) {
  switch (type) {
    case "write_data_to_excel": {
      const { start_cell, data } = params as {
        start_cell: { sheet: string; cell: string };
        data: (string | number | boolean | null)[][];
      };
      const sheet = getSheetByName(spread, start_cell.sheet);
      if (!sheet) throw new Error(`Sheet '${start_cell.sheet}' not found`);
      const { row, col } = parseCellRef(start_cell.cell);
      const cleanData = data.map((r) =>
        r.map((v) => (v === null || v === undefined ? "" : v))
      );
      sheet.setArray(row, col, cleanData);
      break;
    }

    case "write_query_results": {
      const { sheet_name, start_cell, columns, rows } = params as {
        sheet_name: string;
        start_cell: string;
        columns: string[];
        rows: Record<string, string | number | boolean | null>[];
      };
      if (!sheet_name) {
        throw new Error("write_query_results requires sheet_name");
      }
      if (!start_cell) {
        throw new Error("write_query_results requires start_cell");
      }

      const sheet = ensureSheet(spread, sheet_name);
      const { row, col } = parseCellRef(start_cell);
      const tableData = [
        columns,
        ...rows.map((item) =>
          columns.map((column) => {
            const value = item[column];
            return value === null || value === undefined ? "" : value;
          })
        ),
      ];

      sheet.setArray(row, col, tableData);

      const headerRange = sheet.getRange(row, col, 1, columns.length);
      const headerStyle = new GC.Spread.Sheets.Style();
      headerStyle.fontWeight = "700";
      headerStyle.backColor = "#d3d3d3";
      headerStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
      headerRange.setStyle(headerStyle);

      for (let columnIndex = 0; columnIndex < columns.length; columnIndex += 1) {
        sheet.autoFitColumn(col + columnIndex);
      }

      if (rows.length > 0) {
        const tableName = `QueryTable_${Date.now()}`;
        sheet.tables.add(
          tableName,
          row,
          col,
          rows.length + 1,
          columns.length,
          GC.Spread.Sheets.Tables.TableThemes.medium2
        );
      }

      break;
    }

    case "apply_formula": {
      const { cell, formula } = params as {
        cell: { sheet: string; cell: string };
        formula: string;
      };
      const sheet = getSheetByName(spread, cell.sheet);
      if (!sheet) throw new Error(`Sheet '${cell.sheet}' not found`);
      const { row, col } = parseCellRef(cell.cell);
      sheet.setFormula(row, col, formula);
      break;
    }

    case "format_range": {
      const {
        cell_range,
        bold,
        italic,
        underline,
        font_color,
        bg_color,
        border_style,
        border_color,
        number_format,
        alignment,
        wrap_text,
        merge_cells,
      } =
        params as {
          cell_range: { sheet: string; range: string };
          bold?: boolean;
          italic?: boolean;
          underline?: boolean;
          font_color?: string;
          bg_color?: string;
          border_style?: string;
          border_color?: string;
          number_format?: string;
          alignment?: string;
          wrap_text?: boolean;
          merge_cells?: boolean;
        };
      const sheet = getSheetByName(spread, cell_range.sheet);
      if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);
      const parsed = parseRangeRef(cell_range.range);
      const rowCount = parsed.endRow - parsed.startRow + 1;
      const colCount = parsed.endCol - parsed.startCol + 1;
      const range = sheet.getRange(
        parsed.startRow,
        parsed.startCol,
        rowCount,
        colCount,
        GC.Spread.Sheets.SheetArea.viewport
      );
      const style = new GC.Spread.Sheets.Style();

      if (bold !== undefined) style.fontWeight = bold ? "700" : "400";
      if (italic !== undefined) style.fontStyle = italic ? "italic" : "normal";
      if (underline !== undefined) {
        style.textDecoration = underline
          ? GC.Spread.Sheets.TextDecorationType.underline
          : GC.Spread.Sheets.TextDecorationType.none;
      }
      if (font_color) style.foreColor = argbToHex(font_color);
      if (bg_color) style.backColor = argbToHex(bg_color);
      if (number_format) style.formatter = number_format;
      if (wrap_text !== undefined) style.wordWrap = wrap_text;

      if (alignment) {
        style.hAlign = alignmentToSpread(alignment);
      }
      range.setStyle(style);

      if (border_style) {
        const border = new GC.Spread.Sheets.LineBorder(
          argbToHex(border_color ?? "FF000000"),
          borderStyleToSpread(border_style)
        );
        range.setBorder(border, { all: true });
      }

      if (merge_cells) {
        sheet.addSpan(
          parsed.startRow,
          parsed.startCol,
          rowCount,
          colCount,
          GC.Spread.Sheets.SheetArea.viewport
        );
      }

      break;
    }

    case "create_worksheet": {
      const { sheet_name } = params as { sheet_name: string };
      if (!sheet_name) {
        throw new Error("create_worksheet requires sheet_name");
      }
      const firstSheet = spread.getSheet(0);
      if (
        spread.getSheetCount() === 1 &&
        firstSheet &&
        firstSheet.name() === "Sheet1" &&
        isSheetBlank(firstSheet)
      ) {
        firstSheet.name(sheet_name);
        normalizeSheet(firstSheet);
        break;
      }
      const sheet = new GC.Spread.Sheets.Worksheet(sheet_name);
      normalizeSheet(sheet);
      spread.addSheet(
        spread.getSheetCount(),
        sheet
      );
      break;
    }

    case "delete_worksheet": {
      const { sheet_name } = params as { sheet_name: string };
      if (!sheet_name) {
        throw new Error("delete_worksheet requires sheet_name");
      }
      const sheetIndex = getSheetIndexByName(spread, sheet_name);
      if (sheetIndex >= 0 && spread.getSheetCount() > 1) {
        spread.removeSheet(sheetIndex);
      }
      break;
    }

    case "rename_worksheet": {
      const { old_name, new_name } = params as {
        old_name: string;
        new_name: string;
      };
      if (!old_name || !new_name) {
        throw new Error("rename_worksheet requires old_name and new_name");
      }
      const sheet = getSheetByName(spread, old_name);
      if (sheet) sheet.name(new_name);
      break;
    }

    case "clear_cells": {
      const { cell_range } = params as {
        cell_range: { sheet: string; range: string };
      };
      const sheet = getSheetByName(spread, cell_range.sheet);
      if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);
      const parsed = parseRangeRef(cell_range.range);
      sheet.clear(
        parsed.startRow,
        parsed.startCol,
        parsed.endRow - parsed.startRow + 1,
        parsed.endCol - parsed.startCol + 1,
        GC.Spread.Sheets.SheetArea.viewport,
        GC.Spread.Sheets.StorageType.data | GC.Spread.Sheets.StorageType.style
      );
      break;
    }

    case "merge_cells": {
      const { cell_range } = params as {
        cell_range: { sheet: string; range: string };
      };
      const sheet = getSheetByName(spread, cell_range.sheet);
      if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);
      const parsed = parseRangeRef(cell_range.range);
      sheet.addSpan(
        parsed.startRow,
        parsed.startCol,
        parsed.endRow - parsed.startRow + 1,
        parsed.endCol - parsed.startCol + 1,
        GC.Spread.Sheets.SheetArea.viewport
      );
      break;
    }

    case "sort_range": {
      const { cell_range, sort_by_column, ascending } = params as {
        cell_range: { sheet: string; range: string };
        sort_by_column: string;
        ascending?: boolean;
      };
      const sheet = getSheetByName(spread, cell_range.sheet);
      if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);
      const parsed = parseRangeRef(cell_range.range);
      sheet.sortRange(
        parsed.startRow,
        parsed.startCol,
        parsed.endRow - parsed.startRow + 1,
        parsed.endCol - parsed.startCol + 1,
        true,
        [
          {
            index: columnNameToIndex(sort_by_column),
            ascending: ascending ?? true,
          },
        ]
      );
      break;
    }

    case "read_ranges":
    case "write_df_to_sheet":
    case "create_pivot_table":
      console.log(`Operation '${type}' is handled elsewhere or not yet implemented`, params);
      break;

    case "conditional_format_range":
      applyConditionalFormatting(spread, params);
      break;

    case "format_worksheet":
      if (!(params as { sheet_name?: string }).sheet_name) {
        throw new Error("format_worksheet requires sheet_name");
      }
      applyWorksheetFormatting(spread, params);
      break;

    case "copy_worksheet": {
      const { source_sheet, target_sheet } = params as {
        source_sheet: string;
        target_sheet: string;
      };
      const sourceSheet = getSheetByName(spread, source_sheet);
      if (!sourceSheet) throw new Error(`Sheet '${source_sheet}' not found`);
      const clonedSheet = new GC.Spread.Sheets.Worksheet(target_sheet);
      clonedSheet.fromJSON(sourceSheet.toJSON());
      clonedSheet.name(target_sheet);
      spread.addSheet(spread.getSheetCount(), clonedSheet);
      break;
    }

    case "create_table":
      createNativeTable(spread, params);
      break;

    case "drag_formula":
      dragFormula(spread, params);
      break;

    case "copy_range":
      copyRange(spread, params);
      break;

    case "delete_range":
      deleteRange(spread, params);
      break;

    case "unmerge_cells":
      unmergeRange(spread, params);
      break;

    case "set_cell_dropdown":
      setDropdown(spread, params);
      break;

    case "create_chart":
      createChart(spread, params);
      break;

    default:
      console.warn(`Unknown operation type: ${type}`);
  }
}

function normalizeSheet(sheet: GC.Spread.Sheets.Worksheet) {
  if (sheet.getColumnCount() < MIN_SHEET_COLUMN_COUNT) {
    sheet.setColumnCount(MIN_SHEET_COLUMN_COUNT);
  }
  if (sheet.getRowCount() < MIN_SHEET_ROW_COUNT) {
    sheet.setRowCount(MIN_SHEET_ROW_COUNT);
  }

  sheet.options.gridline = {
    color: "#e7ddcf",
    showHorizontalGridline: true,
    showVerticalGridline: true,
  };
}

/** Parse a cell reference like "A1" to 0-indexed {row, col}. */
function parseCellRef(ref: string): { row: number; col: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error(`Invalid cell reference: ${ref}`);
  const colStr = match[1].toUpperCase();
  const row = parseInt(match[2]) - 1;
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { row, col: col - 1 };
}

/** Parse a range like "A1:D10" to 0-indexed bounds. */
function parseRangeRef(range: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} {
  const parts = range.split(":");
  const start = parseCellRef(parts[0]);
  const end = parts.length > 1 ? parseCellRef(parts[1]) : start;
  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

/** Convert 8-char ARGB hex (e.g. "FF0000FF") to CSS hex (#0000FF). */
function argbToHex(argb: string): string {
  if (argb.startsWith("#")) return argb;
  return `#${argb.slice(2)}`;
}

function getSheetByName(spread: SpreadsheetApi, name: string) {
  for (let i = 0; i < spread.getSheetCount(); i++) {
    const sheet = spread.getSheet(i);
    if (sheet?.name() === name) return sheet;
  }
  return null;
}

function ensureSheet(spread: SpreadsheetApi, name: string) {
  const existing = getSheetByName(spread, name);
  if (existing) {
    return existing;
  }

  const firstSheet = spread.getSheet(0);
  if (
    spread.getSheetCount() === 1 &&
    firstSheet &&
    firstSheet.name() === "Sheet1" &&
    isSheetBlank(firstSheet)
  ) {
    firstSheet.name(name);
    normalizeSheet(firstSheet);
    return firstSheet;
  }

  const sheet = new GC.Spread.Sheets.Worksheet(name);
  normalizeSheet(sheet);
  spread.addSheet(spread.getSheetCount(), sheet);
  return sheet;
}

function isSheetBlank(sheet: GC.Spread.Sheets.Worksheet) {
  const usedRange = sheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.all);
  if (!usedRange) return true;

  if (usedRange.rowCount === 1 && usedRange.colCount === 1) {
    const value = sheet.getValue(usedRange.row, usedRange.col);
    const formula = sheet.getFormula(usedRange.row, usedRange.col);
    return (value === null || value === "" || value === undefined) && !formula;
  }

  return false;
}

function getSheetIndexByName(spread: SpreadsheetApi, name: string) {
  for (let i = 0; i < spread.getSheetCount(); i++) {
    const sheet = spread.getSheet(i);
    if (sheet?.name() === name) return i;
  }
  return -1;
}

function alignmentToSpread(alignment: string) {
  switch (alignment) {
    case "center":
      return GC.Spread.Sheets.HorizontalAlign.center;
    case "right":
      return GC.Spread.Sheets.HorizontalAlign.right;
    case "left":
    default:
      return GC.Spread.Sheets.HorizontalAlign.left;
  }
}

function borderStyleToSpread(style: string) {
  switch (style) {
    case "double":
      return GC.Spread.Sheets.LineStyle.double;
    case "thick":
      return GC.Spread.Sheets.LineStyle.thick;
    case "medium":
      return GC.Spread.Sheets.LineStyle.medium;
    case "thin":
    default:
      return GC.Spread.Sheets.LineStyle.thin;
  }
}

function toSpreadRange(range: { startRow: number; startCol: number; endRow: number; endCol: number }) {
  return new GC.Spread.Sheets.Range(
    range.startRow,
    range.startCol,
    range.endRow - range.startRow + 1,
    range.endCol - range.startCol + 1
  );
}

function createConditionalStyle(params: Record<string, unknown>) {
  const style = new GC.Spread.Sheets.Style();
  const bgColor = typeof params.bg_color === "string" ? params.bg_color : undefined;
  const fontColor = typeof params.font_color === "string" ? params.font_color : undefined;

  style.backColor = argbToHex(bgColor ?? "FFFFFFCC");
  if (fontColor) style.foreColor = argbToHex(fontColor);
  return style;
}

function applyConditionalFormatting(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { cell_range, rule_type, rule_params } = params as {
    cell_range: { sheet: string; range: string };
    rule_type: string;
    rule_params: Record<string, string | number | boolean>;
  };
  const sheet = getSheetByName(spread, cell_range.sheet);
  if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);

  const parsedRange = parseRangeRef(cell_range.range);
  const ranges = [toSpreadRange(parsedRange)];
  const style = createConditionalStyle(rule_params);

  switch (rule_type) {
    case "color_scale":
      sheet.conditionalFormats.add3ScaleRule(
        GC.Spread.Sheets.ConditionalFormatting.ScaleValueType.lowestValue,
        Number(rule_params.min_value ?? 0),
        String(rule_params.min_color ?? "#f8d6cc"),
        GC.Spread.Sheets.ConditionalFormatting.ScaleValueType.percentile,
        Number(rule_params.mid_value ?? 50),
        String(rule_params.mid_color ?? "#f3eed5"),
        GC.Spread.Sheets.ConditionalFormatting.ScaleValueType.highestValue,
        Number(rule_params.max_value ?? 100),
        String(rule_params.max_color ?? "#9cb8a4"),
        ranges
      );
      break;

    case "data_bar":
      sheet.conditionalFormats.addDataBarRule(
        GC.Spread.Sheets.ConditionalFormatting.ScaleValueType.lowestValue,
        Number(rule_params.min_value ?? 0),
        GC.Spread.Sheets.ConditionalFormatting.ScaleValueType.highestValue,
        Number(rule_params.max_value ?? 100),
        String(rule_params.color ?? "#8f6b4c"),
        ranges
      );
      break;

    case "formula":
      sheet.conditionalFormats.addFormulaRule(
        String(rule_params.formula ?? ""),
        style,
        ranges
      );
      break;

    case "cell_is":
    default:
      sheet.conditionalFormats.addCellValueRule(
        comparisonOperatorToSpread(String(rule_params.operator ?? "greaterThan")),
        rule_params.value1 ?? 0,
        rule_params.value2 ?? 0,
        style,
        ranges
      );
      break;
  }
}

function comparisonOperatorToSpread(operator: string) {
  const cf = GC.Spread.Sheets.ConditionalFormatting.ComparisonOperators;
  switch (operator) {
    case "equals":
    case "equal":
    case "equalsTo":
      return cf.equalsTo;
    case "notEquals":
    case "notEqualsTo":
      return cf.notEqualsTo;
    case "greaterThanOrEqual":
    case "greaterThanOrEqualsTo":
      return cf.greaterThanOrEqualsTo;
    case "lessThan":
      return cf.lessThan;
    case "lessThanOrEqual":
    case "lessThanOrEqualsTo":
      return cf.lessThanOrEqualsTo;
    case "between":
      return cf.between;
    case "notBetween":
      return cf.notBetween;
    case "greaterThan":
    default:
      return cf.greaterThan;
  }
}

function applyWorksheetFormatting(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { sheet_name, show_gridlines, tab_color } = params as {
    sheet_name: string;
    show_gridlines?: boolean;
    tab_color?: string;
  };
  const sheet = getSheetByName(spread, sheet_name);
  if (!sheet) throw new Error(`Sheet '${sheet_name}' not found`);

  if (show_gridlines !== undefined) {
    sheet.options.gridline = {
      ...(sheet.options.gridline ?? {}),
      showHorizontalGridline: show_gridlines,
      showVerticalGridline: show_gridlines,
    };
  }

  if (tab_color) {
    sheet.options.sheetTabColor = argbToHex(tab_color);
  }
}

function createNativeTable(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { data_range, table_name, table_style } = params as {
    data_range: { sheet: string; range: string };
    table_name: string;
    table_style?: string;
  };
  const sheet = getSheetByName(spread, data_range.sheet);
  if (!sheet) throw new Error(`Sheet '${data_range.sheet}' not found`);

  const parsed = parseRangeRef(data_range.range);
  const themeMap = GC.Spread.Sheets.Tables.TableThemes as unknown as Record<
    string,
    GC.Spread.Sheets.Tables.TableTheme
  >;

  sheet.tables.add(
    table_name,
    parsed.startRow,
    parsed.startCol,
    parsed.endRow - parsed.startRow + 1,
    parsed.endCol - parsed.startCol + 1,
    themeMap[table_style ?? "medium2"] ?? GC.Spread.Sheets.Tables.TableThemes.medium2
  );
}

function dragFormula(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { selected_range, drag_direction, drag_distance } = params as {
    selected_range: { sheet: string; range: string };
    drag_direction: "down" | "up" | "left" | "right";
    drag_distance: number;
  };
  const sheet = getSheetByName(spread, selected_range.sheet);
  if (!sheet) throw new Error(`Sheet '${selected_range.sheet}' not found`);

  const parsed = parseRangeRef(selected_range.range);
  const rowCount = parsed.endRow - parsed.startRow + 1;
  const colCount = parsed.endCol - parsed.startCol + 1;
  const option = GC.Spread.Sheets.CopyToOptions.formula | GC.Spread.Sheets.CopyToOptions.style;

  switch (drag_direction) {
    case "up":
      for (let step = 1; step <= drag_distance; step++) {
        sheet.copyTo(
          parsed.startRow,
          parsed.startCol,
          parsed.startRow - step,
          parsed.startCol,
          rowCount,
          colCount,
          option
        );
      }
      break;
    case "left":
      for (let step = 1; step <= drag_distance; step++) {
        sheet.copyTo(
          parsed.startRow,
          parsed.startCol,
          parsed.startRow,
          parsed.startCol - step,
          rowCount,
          colCount,
          option
        );
      }
      break;
    case "right":
      for (let step = 1; step <= drag_distance; step++) {
        sheet.copyTo(
          parsed.startRow,
          parsed.startCol,
          parsed.startRow,
          parsed.endCol + step,
          rowCount,
          colCount,
          option
        );
      }
      break;
    case "down":
    default:
      for (let step = 1; step <= drag_distance; step++) {
        sheet.copyTo(
          parsed.startRow,
          parsed.startCol,
          parsed.endRow + step,
          parsed.startCol,
          rowCount,
          colCount,
          option
        );
      }
      break;
  }
}

function copyRange(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { source_range, target_cell } = params as {
    source_range: { sheet: string; range: string };
    target_cell: { sheet: string; cell: string };
  };
  if (source_range.sheet !== target_cell.sheet) {
    throw new Error("copy_range currently supports copying within the same sheet");
  }

  const sheet = getSheetByName(spread, source_range.sheet);
  if (!sheet) throw new Error(`Sheet '${source_range.sheet}' not found`);

  const source = parseRangeRef(source_range.range);
  const target = parseCellRef(target_cell.cell);
  sheet.copyTo(
    source.startRow,
    source.startCol,
    target.row,
    target.col,
    source.endRow - source.startRow + 1,
    source.endCol - source.startCol + 1,
    GC.Spread.Sheets.CopyToOptions.all
  );
}

function deleteRange(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { cell_range, shift_direction } = params as {
    cell_range: { sheet: string; range: string };
    shift_direction: "up" | "left";
  };
  const sheet = getSheetByName(spread, cell_range.sheet);
  if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);

  const range = parseRangeRef(cell_range.range);
  const rowCount = range.endRow - range.startRow + 1;
  const colCount = range.endCol - range.startCol + 1;

  if (shift_direction === "up") {
    const totalRows = sheet.getRowCount(GC.Spread.Sheets.SheetArea.viewport);
    const trailingRows = totalRows - range.endRow - 1;
    if (trailingRows > 0) {
      const data = sheet.getArray(
        range.endRow + 1,
        range.startCol,
        trailingRows,
        colCount,
        true
      );
      sheet.setArray(range.startRow, range.startCol, data, true);
    }
    sheet.clear(
      totalRows - rowCount,
      range.startCol,
      rowCount,
      colCount,
      GC.Spread.Sheets.SheetArea.viewport,
      GC.Spread.Sheets.StorageType.data | GC.Spread.Sheets.StorageType.style
    );
    return;
  }

  const totalCols = sheet.getColumnCount(GC.Spread.Sheets.SheetArea.viewport);
  const trailingCols = totalCols - range.endCol - 1;
  if (trailingCols > 0) {
    const data = sheet.getArray(
      range.startRow,
      range.endCol + 1,
      rowCount,
      trailingCols,
      true
    );
    sheet.setArray(range.startRow, range.startCol, data, true);
  }
  sheet.clear(
    range.startRow,
    totalCols - colCount,
    rowCount,
    colCount,
    GC.Spread.Sheets.SheetArea.viewport,
    GC.Spread.Sheets.StorageType.data | GC.Spread.Sheets.StorageType.style
  );
}

function unmergeRange(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { cell_range } = params as {
    cell_range: { sheet: string; range: string };
  };
  const sheet = getSheetByName(spread, cell_range.sheet);
  if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);

  const range = parseRangeRef(cell_range.range);
  const spans = sheet.getSpans(
    new GC.Spread.Sheets.Range(
      range.startRow,
      range.startCol,
      range.endRow - range.startRow + 1,
      range.endCol - range.startCol + 1
    ),
    GC.Spread.Sheets.SheetArea.viewport
  );

  for (const span of spans) {
    sheet.removeSpan(span.row, span.col, GC.Spread.Sheets.SheetArea.viewport);
  }
}

function setDropdown(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { cell_range, dropdown_params } = params as {
    cell_range: { sheet: string; range: string };
    dropdown_params: { options: string[]; allow_blank?: boolean };
  };
  const sheet = getSheetByName(spread, cell_range.sheet);
  if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);

  const range = parseRangeRef(cell_range.range);
  const validator = GC.Spread.Sheets.DataValidation.createListValidator(
    dropdown_params.options.join(",")
  );
  validator.ignoreBlank(dropdown_params.allow_blank ?? true);
  sheet.setDataValidator(
    range.startRow,
    range.startCol,
    range.endRow - range.startRow + 1,
    range.endCol - range.startCol + 1,
    validator,
    GC.Spread.Sheets.SheetArea.viewport
  );
}

function createChart(
  spread: SpreadsheetApi,
  params: Record<string, unknown>
) {
  const { cell_range, chart_type, chart_name, title, anchor_cell } = params as {
    cell_range: { sheet: string; range: string };
    chart_type: string;
    chart_name?: string;
    title?: string;
    anchor_cell?: string;
  };
  const sheet = getSheetByName(spread, cell_range.sheet);
  if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);

  const anchor = parseCellRef(anchor_cell ?? "H2");
  const chartTypes = GC.Spread.Sheets.Charts.ChartType as unknown as Record<
    string,
    GC.Spread.Sheets.Charts.ChartType
  >;
  const chart = sheet.charts.add(
    chart_name ?? `chart_${Date.now()}`,
    chartTypes[chart_type] ?? GC.Spread.Sheets.Charts.ChartType.columnClustered,
    anchor.col * 88,
    anchor.row * 26,
    520,
    320,
    cell_range.range
  );

  if (title) {
    chart.title({ text: title });
  }
}

function columnNameToIndex(column: string) {
  let index = 0;
  const colStr = column.toUpperCase();
  for (let i = 0; i < colStr.length; i++) {
    index = index * 26 + (colStr.charCodeAt(i) - 64);
  }
  return index - 1;
}
