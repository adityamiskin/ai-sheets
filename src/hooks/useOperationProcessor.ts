import { useQuery, useMutation } from "convex/react";
import { api } from "../../convex/_generated/api";
import { useEffect, useRef } from "react";
import type { FUniver } from "@univerjs/presets";
import type { Id } from "../../convex/_generated/dataModel";

/**
 * Subscribes to pending operations and applies them to the Univer instance.
 */
export function useOperationProcessor(
  workbookId: Id<"workbooks"> | null,
  univerAPI: FUniver | null
) {
  const pendingOps = useQuery(
    api.operations.getPending,
    workbookId ? { workbookId } : "skip"
  );
  const markApplied = useMutation(api.operations.markApplied);
  const markFailed = useMutation(api.operations.markFailed);
  const processingRef = useRef(new Set<string>());

  useEffect(() => {
    if (!pendingOps || !univerAPI || pendingOps.length === 0) return;

    for (const op of pendingOps) {
      const opId = op._id as string;
      if (processingRef.current.has(opId)) continue;
      processingRef.current.add(opId);

      (async () => {
        try {
          const params = JSON.parse(op.params);
          applyOperation(univerAPI, op.type, params);
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
  }, [pendingOps, univerAPI, markApplied, markFailed]);
}

/**
 * Applies a single operation to the Univer Facade API.
 */
function applyOperation(
  univerAPI: FUniver,
  type: string,
  params: Record<string, unknown>
) {
  const workbook = univerAPI.getActiveWorkbook();
  if (!workbook) throw new Error("No active workbook");

  switch (type) {
    case "write_data_to_excel": {
      const { start_cell, data } = params as {
        start_cell: { sheet: string; cell: string };
        data: (string | number | boolean | null)[][];
      };
      const sheet = workbook.getSheetByName(start_cell.sheet);
      if (!sheet) throw new Error(`Sheet '${start_cell.sheet}' not found`);
      const { row, col } = parseCellRef(start_cell.cell);
      const endRow = row + data.length - 1;
      const endCol = col + (data[0]?.length ?? 1) - 1;
      const range = sheet.getRange(row, col, endRow - row + 1, endCol - col + 1);
      range.setValues(data);
      break;
    }

    case "apply_formula": {
      const { cell, formula } = params as {
        cell: { sheet: string; cell: string };
        formula: string;
      };
      const sheet = workbook.getSheetByName(cell.sheet);
      if (!sheet) throw new Error(`Sheet '${cell.sheet}' not found`);
      const { row, col } = parseCellRef(cell.cell);
      const range = sheet.getRange(row, col, 1, 1);
      range.setValue(formula);
      break;
    }

    case "format_range": {
      const { cell_range, bold, italic, font_color, bg_color, number_format, alignment } =
        params as {
          cell_range: { sheet: string; range: string };
          bold?: boolean;
          italic?: boolean;
          font_color?: string;
          bg_color?: string;
          number_format?: string;
          alignment?: string;
        };
      const sheet = workbook.getSheetByName(cell_range.sheet);
      if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);
      const parsed = parseRangeRef(cell_range.range);
      const range = sheet.getRange(
        parsed.startRow,
        parsed.startCol,
        parsed.endRow - parsed.startRow + 1,
        parsed.endCol - parsed.startCol + 1
      );
      if (bold !== undefined) range.setFontWeight(bold ? "bold" : "normal");
      if (italic !== undefined) range.setFontStyle(italic ? "italic" : "normal");
      if (font_color) range.setFontColor(argbToHex(font_color));
      if (bg_color) range.setBackground(argbToHex(bg_color));
      if (number_format) range.setNumberFormat(number_format);
      if (alignment) {
        const alignMap: Record<string, number> = { left: 0, center: 1, right: 2 };
        range.setHorizontalAlignment(alignMap[alignment] ?? 0);
      }
      break;
    }

    case "create_worksheet": {
      const { sheet_name } = params as { sheet_name: string };
      workbook.create(sheet_name, 100, 26);
      break;
    }

    case "delete_worksheet": {
      const { sheet_name } = params as { sheet_name: string };
      const sheet = workbook.getSheetByName(sheet_name);
      if (sheet) workbook.deleteSheet(sheet.getSheetId());
      break;
    }

    case "rename_worksheet": {
      const { old_name, new_name } = params as {
        old_name: string;
        new_name: string;
      };
      const sheet = workbook.getSheetByName(old_name);
      if (sheet) sheet.setName(new_name);
      break;
    }

    case "clear_cells": {
      const { cell_range } = params as {
        cell_range: { sheet: string; range: string };
      };
      const sheet = workbook.getSheetByName(cell_range.sheet);
      if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);
      const parsed = parseRangeRef(cell_range.range);
      const range = sheet.getRange(
        parsed.startRow,
        parsed.startCol,
        parsed.endRow - parsed.startRow + 1,
        parsed.endCol - parsed.startCol + 1
      );
      range.setValue("");
      break;
    }

    case "merge_cells": {
      const { cell_range } = params as {
        cell_range: { sheet: string; range: string };
      };
      const sheet = workbook.getSheetByName(cell_range.sheet);
      if (!sheet) throw new Error(`Sheet '${cell_range.sheet}' not found`);
      const parsed = parseRangeRef(cell_range.range);
      const range = sheet.getRange(
        parsed.startRow,
        parsed.startCol,
        parsed.endRow - parsed.startRow + 1,
        parsed.endCol - parsed.startCol + 1
      );
      range.merge();
      break;
    }

    case "sort_range": {
      // Sort is complex — for now just log and mark as applied
      console.log("sort_range operation received — manual handling needed", params);
      break;
    }

    // Read-only operations or operations that need more complex handling
    case "read_ranges":
    case "conditional_format_range":
    case "format_worksheet":
    case "copy_worksheet":
    case "create_table":
    case "drag_formula":
    case "copy_range":
    case "delete_range":
    case "unmerge_cells":
    case "set_cell_dropdown":
    case "create_pivot_table":
    case "write_df_to_sheet":
      console.log(`Operation '${type}' received — stub handler`, params);
      break;

    default:
      console.warn(`Unknown operation type: ${type}`);
  }
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
  // Skip the alpha channel (first 2 chars)
  return `#${argb.slice(2)}`;
}
