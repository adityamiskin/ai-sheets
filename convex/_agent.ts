"use node";

import { Agent, createTool, type ToolCtx } from "@convex-dev/agent";
import { createOpenAI } from "@ai-sdk/openai";
import { api, components } from "./_generated/api";
import type { DataModel } from "./_generated/dataModel";
import { z } from "zod";
import type { Id } from "./_generated/dataModel";

// --- Custom tool context: workbookId is spread onto ctx from the action ---

type SheetsToolCtx = ToolCtx<DataModel> & { workbookId: Id<"workbooks"> };

/** Extract the workbookId that was spread onto ctx by the calling action. */
function wbId(ctx: ToolCtx): Id<"workbooks"> {
  const id = (ctx as SheetsToolCtx).workbookId;
  if (!id) throw new Error("No workbook context — workbookId not on tool ctx");
  return id;
}

// --- Zod helpers ---

const CellSchema = z.object({
  sheet: z.string().describe("Sheet name"),
  cell: z.string().describe("Cell reference like A1, B2"),
});

const SheetRangeSchema = z.object({
  sheet: z.string().describe("Sheet name"),
  range: z.string().describe("Range like A1:D10"),
});

const RangeItemSchema = z.object({
  sheet: z.string().describe("Sheet name"),
  range: z.string().describe("Range like A1:D10"),
});

// --- System prompt ---

const SYSTEM_PROMPT = `You are Ramp Sheets - a spreadsheet agent created by Ramp Labs. Your primary task is to both help the user understand data in the spreadsheet and operate on their spreadsheet by using tools provided to you.

IMPORTANT RULES:
1. When possible use write_data_to_excel to set multiple cells at once.
2. There is no need to repeat the spreadsheet state in the chat, as the user can always access the spreadsheet.
3. When generating text / markdown output, ONLY generate h5 tags at the highest level (e.g. #####).
4. Before you start working on a spreadsheet, add all tasks you think you'll need using add_tasks tool.
5. Focus on one task at a time. As you complete a task, check it off using complete_task.
6. ALWAYS add a final "Review work" task.
7. Work in small parts: insert data or formulas in chunks, then verify each chunk before moving on.
8. After each significant edit, use read_ranges to check your work is correct before continuing.
9. Apply professional investment banking formatting when creating or modifying spreadsheets.
10. NEVER write formulas that reference entire columns (e.g. A:A). Always specify explicit ranges.
11. If you need to apply a formula to many cells, use apply_formula for the first cell, then drag_formula.
12. Avoid hardcoding values in formulas. Place constants in dedicated cells and reference them.

FORMATTING STANDARDS:
- ALL numeric values: RIGHT-ALIGNED
- Text labels: LEFT-ALIGNED
- Column headers: CENTER-ALIGNED, bold, gray background
- BLUE font (FF0000FF): Hardcoded inputs/assumptions
- BLACK font (FF000000): Formulas and calculated values
- GREEN font (FF008000): Cross-sheet references
- RED font (FFFF0000): Errors, warnings
- Currency: _($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)
- Percentages: 0.00%
- Key metrics: Yellow background (FFFFFFCC), bold`;

// ═══════════════════════════════════════════════════════
// Task Management Tools (4)
// ═══════════════════════════════════════════════════════

const readTasks = createTool({
  description: "Returns a list of todos/tasks for the current workbook.",
  args: z.object({}),
  handler: async (ctx): Promise<{ tasks: Array<{ _id: string; content: string; status: string; order: number }> }> => {
    const tasks = await ctx.runQuery(api.tasks.list, {
      workbookId: wbId(ctx),
    });
    return { tasks: tasks as Array<{ _id: string; content: string; status: string; order: number }> };
  },
});

const addTasks = createTool({
  description: "Adds list of tasks to the task list.",
  args: z.object({
    tasks: z.array(
      z.object({
        content: z.string().describe("Task description"),
      })
    ),
  }),
  handler: async (ctx, args): Promise<{ added: number; taskIds: string[] }> => {
    const ids = await ctx.runMutation(api.tasks.add, {
      workbookId: wbId(ctx),
      threadId: ctx.threadId,
      tasks: args.tasks,
    });
    return { added: ids.length, taskIds: ids as string[] };
  },
});

const completeTask = createTool({
  description: "Marks a task as completed by its number (1-indexed).",
  args: z.object({
    task_number: z.number().describe("The 1-indexed task number to complete"),
  }),
  handler: async (ctx, args) => {
    await ctx.runMutation(api.tasks.complete, {
      workbookId: wbId(ctx),
      taskNumber: args.task_number,
    });
    return { completed: args.task_number };
  },
});

const deleteTasks = createTool({
  description:
    "Deletes specific tasks or all tasks from the task list. Pass null for task_numbers to delete all.",
  args: z.object({
    task_numbers: z
      .array(z.number())
      .nullable()
      .describe("Array of 1-indexed task numbers to delete, or null to delete all"),
  }),
  handler: async (ctx, args) => {
    if (args.task_numbers === null) {
      await ctx.runMutation(api.tasks.deleteAll, {
        workbookId: wbId(ctx),
      });
      return { deleted: "all" };
    } else {
      await ctx.runMutation(api.tasks.deleteSpecific, {
        workbookId: wbId(ctx),
        taskNumbers: args.task_numbers,
      });
      return { deleted: args.task_numbers };
    }
  },
});

// ═══════════════════════════════════════════════════════
// Spreadsheet Reading Tools (7)
// ═══════════════════════════════════════════════════════

const readRanges = createTool({
  description:
    "Read data from Excel worksheet with cell metadata including validation rules. Returns cell values, formulas, and formatting info.",
  args: z.object({
    ranges: z.array(RangeItemSchema).describe("Array of sheet ranges to read"),
  }),
  handler: async (ctx, args): Promise<Record<string, RangeReadResult | { error: string }>> => {
    const workbookId = wbId(ctx);
    const workbook = await ctx.runQuery(api.workbooks.get, { id: workbookId }) as { snapshot: string } | null;
    if (!workbook) return { _error: { error: "Workbook not found" } };

    const snapshot = JSON.parse(workbook.snapshot) as WorkbookSnapshot;
    const results: Record<string, RangeReadResult | { error: string }> = {};

    for (const range of args.ranges) {
      const sheetData = findSheet(snapshot, range.sheet);
      if (!sheetData) {
        results[`${range.sheet}!${range.range}`] = {
          error: `Sheet '${range.sheet}' not found`,
        };
        continue;
      }

      const parsed = parseRange(range.range);
      const cells: CellReadValue[][] = [];

      for (let r = parsed.startRow; r <= parsed.endRow; r++) {
        const row: CellReadValue[] = [];
        for (let c = parsed.startCol; c <= parsed.endCol; c++) {
          const cellData = sheetData.cellData?.[r]?.[c];
          row.push(
            cellData
              ? { value: cellData.v, formula: cellData.f, type: cellData.t }
              : null
          );
        }
        cells.push(row);
      }

      results[`${range.sheet}!${range.range}`] = {
        cells,
        rowCount: parsed.endRow - parsed.startRow + 1,
        colCount: parsed.endCol - parsed.startCol + 1,
      };
    }

    // Store as a read-only operation for the activity feed
    await ctx.runMutation(api.operations.createWithResult, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "read_ranges",
      params: JSON.stringify(args),
      result: JSON.stringify(results),
      explanation: `Read ${args.ranges.map((r) => `${r.sheet}!${r.range}`).join(", ")}`,
    });

    return results;
  },
});

const getWorkbookMetadata = createTool({
  description:
    "Get metadata about workbook including sheet names, tab colors, and used ranges.",
  args: z.object({
    include_ranges: z.boolean().default(true).describe("Whether to include used ranges"),
  }),
  handler: async (ctx, args): Promise<{ name: string; sheets: Array<{ name: string; id: string; tabColor?: string; rowCount?: number; columnCount?: number; usedRange?: string | null }>; sheetCount: number } | { error: string }> => {
    const workbookId = wbId(ctx);
    const workbook = await ctx.runQuery(api.workbooks.get, { id: workbookId }) as { name: string; snapshot: string } | null;
    if (!workbook) return { error: "Workbook not found" };

    const snapshot = JSON.parse(workbook.snapshot) as WorkbookSnapshot;
    const sheets = Object.values(snapshot.sheets).map((s) => ({
      name: s.name,
      id: s.id,
      tabColor: s.tabColor,
      rowCount: s.rowCount,
      columnCount: s.columnCount,
      usedRange: args.include_ranges ? getUsedRange(s.cellData) : undefined,
    }));

    return {
      name: workbook.name,
      sheets,
      sheetCount: sheets.length,
    };
  },
});

const getDataValidationInfo = createTool({
  description: "Get all data validation rules in a worksheet.",
  args: z.object({
    sheet_name: z.string().describe("Name of the sheet"),
  }),
  handler: async (ctx, args): Promise<{ validations: unknown[] } | { error: string }> => {
    const workbookId = wbId(ctx);
    const workbook = await ctx.runQuery(api.workbooks.get, { id: workbookId }) as { snapshot: string } | null;
    if (!workbook) return { error: "Workbook not found" };

    const snapshot = JSON.parse(workbook.snapshot) as WorkbookSnapshot;
    const sheet = findSheet(snapshot, args.sheet_name);
    if (!sheet) return { error: `Sheet '${args.sheet_name}' not found` };

    return { validations: sheet.dataValidations ?? [] };
  },
});

const getMergedCells = createTool({
  description: "Get merged cells in a worksheet.",
  args: z.object({
    sheet_name: z.string().describe("Name of the sheet"),
  }),
  handler: async (ctx, args): Promise<{ mergedCells: unknown[] } | { error: string }> => {
    const workbookId = wbId(ctx);
    const workbook = await ctx.runQuery(api.workbooks.get, { id: workbookId }) as { snapshot: string } | null;
    if (!workbook) return { error: "Workbook not found" };

    const snapshot = JSON.parse(workbook.snapshot) as WorkbookSnapshot;
    const sheet = findSheet(snapshot, args.sheet_name);
    if (!sheet) return { error: `Sheet '${args.sheet_name}' not found` };

    return { mergedCells: sheet.mergeData ?? [] };
  },
});

const validateExcelRange = createTool({
  description: "Validate if a range exists and is properly formatted.",
  args: z.object({
    sheet_name: z.string(),
    range: z.string().describe("Range like A1:D10"),
  }),
  handler: async (_ctx, args) => {
    try {
      const parsed = parseRange(args.range);
      return {
        valid: true as const,
        parsed: {
          startRow: parsed.startRow,
          startCol: parsed.startCol,
          endRow: parsed.endRow,
          endCol: parsed.endCol,
        },
      };
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Invalid range";
      return { valid: false as const, error: msg };
    }
  },
});

const validateFormulaSyntax = createTool({
  description: "Validate Excel formula syntax without applying it.",
  args: z.object({
    cell: CellSchema,
    formula: z.string().describe("The formula string starting with ="),
  }),
  handler: async (_ctx, args) => {
    if (!args.formula.startsWith("=")) {
      return { valid: false as const, error: "Formula must start with =" };
    }
    let depth = 0;
    for (const ch of args.formula) {
      if (ch === "(") depth++;
      if (ch === ")") depth--;
      if (depth < 0) return { valid: false as const, error: "Unbalanced parentheses" };
    }
    if (depth !== 0) return { valid: false as const, error: "Unbalanced parentheses" };
    return { valid: true as const, formula: args.formula };
  },
});

const grepTool = createTool({
  description:
    "Search a worksheet for rows that contain the query substring in any cell.",
  args: z.object({
    sheet: z.string().describe("Sheet name"),
    query: z.string().describe("Search query"),
    case_sensitive: z.boolean().default(false),
    include_headers: z.boolean().default(true),
  }),
  handler: async (ctx, args): Promise<{ matches: number; rows: Array<{ row: number; data: Record<string, CellData> }> } | { error: string }> => {
    const workbookId = wbId(ctx);
    const workbook = await ctx.runQuery(api.workbooks.get, { id: workbookId }) as { snapshot: string } | null;
    if (!workbook) return { error: "Workbook not found" };

    const snapshot = JSON.parse(workbook.snapshot) as WorkbookSnapshot;
    const sheet = findSheet(snapshot, args.sheet);
    if (!sheet) return { error: `Sheet '${args.sheet}' not found` };

    const matchingRows: Array<{ row: number; data: Record<string, CellData> }> = [];
    const cellData = sheet.cellData ?? {};
    const searchQuery = args.case_sensitive ? args.query : args.query.toLowerCase();

    for (const [rowStr, rowData] of Object.entries(cellData)) {
      const row = parseInt(rowStr);
      for (const [, cell] of Object.entries(rowData)) {
        const val = String(cell.v ?? "");
        const compareVal = args.case_sensitive ? val : val.toLowerCase();
        if (compareVal.includes(searchQuery)) {
          matchingRows.push({ row: row + 1, data: rowData });
          break;
        }
      }
    }

    return {
      matches: matchingRows.length,
      rows: matchingRows.slice(0, 100),
    };
  },
});

// ═══════════════════════════════════════════════════════
// Spreadsheet Writing Tools (2)
// ═══════════════════════════════════════════════════════

const writeDataToExcel = createTool({
  description:
    "Write tabular data to the specified worksheet cell. Data is a 2D array.",
  args: z.object({
    start_cell: CellSchema,
    data: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))).describe(
      "2D array of values to write"
    ),
    explanation: z.string().describe("Human-readable explanation of what is being written"),
  }),
  handler: async (ctx, args): Promise<{ success: true; cell: string; rows: number; cols: number }> => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "write_data_to_excel",
      params: JSON.stringify({
        start_cell: args.start_cell,
        data: args.data,
      }),
      explanation: args.explanation,
    });
    return {
      success: true,
      cell: `${args.start_cell.sheet}!${args.start_cell.cell}`,
      rows: args.data.length,
      cols: args.data[0]?.length ?? 0,
    };
  },
});

const writeDfToSheet = createTool({
  description:
    "Write a pandas DataFrame to a workbook sheet (from execute_code results).",
  args: z.object({
    df_name: z.string().describe("Name of the DataFrame variable"),
    sheet_name: z.string(),
    include_header: z.boolean().default(true),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "write_df_to_sheet",
      params: JSON.stringify(args),
      explanation: `Write DataFrame '${args.df_name}' to sheet '${args.sheet_name}'`,
    });
    return { success: true };
  },
});

// ═══════════════════════════════════════════════════════
// Formatting Tools (3)
// ═══════════════════════════════════════════════════════

const formatRange = createTool({
  description:
    "Apply formatting to a range of cells. Colors are 8-character aRGB hex without # prefix.",
  args: z.object({
    cell_range: SheetRangeSchema,
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    font_color: z.string().optional().describe("8-char aRGB hex, e.g. FF0000FF"),
    bg_color: z.string().optional().describe("8-char aRGB hex, e.g. FFFFFFCC"),
    border_style: z.string().optional().describe("thin, medium, thick, double"),
    border_color: z.string().optional(),
    number_format: z.string().optional(),
    alignment: z.string().optional().describe("left, center, right"),
    wrap_text: z.boolean().optional(),
    merge_cells: z.boolean().optional(),
    explanation: z.string().describe("What formatting is being applied and why"),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    const { cell_range, explanation, ...formatOpts } = args;
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "format_range",
      params: JSON.stringify({ cell_range, ...formatOpts }),
      explanation,
    });
    return { success: true, range: `${cell_range.sheet}!${cell_range.range}` };
  },
});

const conditionalFormatRange = createTool({
  description:
    "Apply conditional formatting to a range of cells (color scales, data bars, cell conditions, formulas).",
  args: z.object({
    cell_range: SheetRangeSchema,
    rule_type: z.string().describe("Type: color_scale, data_bar, cell_is, formula"),
    rule_params: z
      .record(z.string(), z.union([z.string(), z.number(), z.boolean()]))
      .describe("Rule-specific parameters"),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "conditional_format_range",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const formatWorksheet = createTool({
  description: "Apply sheet-level formatting settings to a worksheet.",
  args: z.object({
    sheet_name: z.string(),
    show_gridlines: z.boolean().optional(),
    tab_color: z.string().optional().describe("8-char aRGB hex"),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "format_worksheet",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

// ═══════════════════════════════════════════════════════
// Sheet Management Tools (5)
// ═══════════════════════════════════════════════════════

const renameWorksheet = createTool({
  description: "Rename worksheet in workbook.",
  args: z.object({
    old_name: z.string(),
    new_name: z.string(),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "rename_worksheet",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const createWorksheet = createTool({
  description: "Create new worksheet in workbook.",
  args: z.object({
    sheet_name: z.string(),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "create_worksheet",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const deleteWorksheet = createTool({
  description: "Delete worksheet from workbook.",
  args: z.object({
    sheet_name: z.string(),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "delete_worksheet",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const copyWorksheet = createTool({
  description: "Copy worksheet within workbook.",
  args: z.object({
    source_sheet: z.string(),
    target_sheet: z.string(),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "copy_worksheet",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const createTable = createTool({
  description: "Creates a native Excel table from a specified range of data.",
  args: z.object({
    data_range: SheetRangeSchema,
    table_name: z.string(),
    table_style: z.string().optional().describe("Table style like TableStyleMedium2"),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "create_table",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

// ═══════════════════════════════════════════════════════
// Cell Operations Tools (9)
// ═══════════════════════════════════════════════════════

const applyFormula = createTool({
  description: "Apply Excel formula to a cell with verification.",
  args: z.object({
    cell: CellSchema,
    formula: z.string().describe("Excel formula starting with ="),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "apply_formula",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true, cell: `${args.cell.sheet}!${args.cell.cell}` };
  },
});

const dragFormula = createTool({
  description:
    "Simulate Excel's formula drag functionality. Drags formulas with auto-increment of relative references.",
  args: z.object({
    selected_range: SheetRangeSchema,
    drag_direction: z.enum(["down", "up", "left", "right"]),
    drag_distance: z.number().describe("Number of cells to drag"),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "drag_formula",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const copyRange = createTool({
  description: "Copy a range of cells to another location.",
  args: z.object({
    source_range: SheetRangeSchema,
    target_cell: CellSchema,
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "copy_range",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const deleteRange = createTool({
  description:
    "Delete a range of cells and shift remaining cells to fill the gap.",
  args: z.object({
    cell_range: SheetRangeSchema,
    shift_direction: z
      .enum(["up", "left"])
      .describe("Direction to shift remaining cells"),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "delete_range",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const clearCells = createTool({
  description:
    "Clear the contents and formatting of a range of cells WITHOUT shifting other cells.",
  args: z.object({
    cell_range: SheetRangeSchema,
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "clear_cells",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const mergeCells = createTool({
  description: "Merge a range of cells.",
  args: z.object({
    cell_range: SheetRangeSchema,
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "merge_cells",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const unmergeCells = createTool({
  description: "Unmerge a range of cells.",
  args: z.object({
    cell_range: SheetRangeSchema,
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "unmerge_cells",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const setCellDropdown = createTool({
  description: "Add a dropdown list to cells for easy data entry validation.",
  args: z.object({
    cell_range: SheetRangeSchema,
    dropdown_params: z.object({
      options: z.array(z.string()).describe("Dropdown options"),
      allow_blank: z.boolean().optional(),
    }),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "set_cell_dropdown",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

const sortRange = createTool({
  description: "Sort a range of cells by a specified column.",
  args: z.object({
    cell_range: SheetRangeSchema,
    sort_by_column: z.string().describe("Column letter to sort by"),
    ascending: z.boolean().default(true),
    has_header: z.boolean().default(true),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "sort_range",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

// ═══════════════════════════════════════════════════════
// Data Operations (1)
// ═══════════════════════════════════════════════════════

const createPivotTable = createTool({
  description: "Create pivot table in worksheet.",
  args: z.object({
    data_range: SheetRangeSchema,
    rows: z.array(z.string()).describe("Row field names"),
    values: z.array(z.string()).describe("Value field names"),
    columns: z.array(z.string()).optional(),
    agg_func: z
      .string()
      .default("sum")
      .describe("Aggregation function: sum, count, average, min, max"),
    explanation: z.string(),
  }),
  handler: async (ctx, args) => {
    const workbookId = wbId(ctx);
    await ctx.runMutation(api.operations.create, {
      workbookId,
      threadId: ctx.threadId ?? "",
      type: "create_pivot_table",
      params: JSON.stringify(args),
      explanation: args.explanation,
    });
    return { success: true };
  },
});

// ═══════════════════════════════════════════════════════
// Web & Research Tools (3)
// ═══════════════════════════════════════════════════════

const searchWeb = createTool({
  description:
    "Perform a web search to gather external context, definitions, benchmarks, or up-to-date information.",
  args: z.object({
    query: z.string().describe("Search query"),
  }),
  handler: async (_ctx, args) => {
    return {
      note: "Web search functionality — use the LLM's built-in web search capability",
      query: args.query,
    };
  },
});

const generateResearchReport = createTool({
  description:
    "Generate a comprehensive research report on a given topic using multiple web searches.",
  args: z.object({
    topic: z.string(),
    research_steps: z.array(z.string()).describe("Steps to research"),
    depth_mode: z.string().default("standard").describe("standard, deep, or quick"),
  }),
  handler: async (_ctx, args) => {
    return {
      note: "Research report generation",
      topic: args.topic,
      steps: args.research_steps,
    };
  },
});

const browseUrl = createTool({
  description: "Browse one or more URLs and extract/analyze their content.",
  args: z.object({
    urls: z.array(z.string()).describe("URLs to browse"),
    query: z.string().describe("What to look for"),
  }),
  handler: async (_ctx, args) => {
    const results: Record<string, { status: number; contentLength: number; preview: string } | { error: string }> =
      {};
    for (const url of args.urls.slice(0, 5)) {
      try {
        const response = await fetch(url);
        const text = await response.text();
        results[url] = {
          status: response.status,
          contentLength: text.length,
          preview: text.substring(0, 2000),
        };
      } catch (e) {
        const msg = e instanceof Error ? e.message : "Fetch failed";
        results[url] = { error: msg };
      }
    }
    return results;
  },
});

// ═══════════════════════════════════════════════════════
// Code Execution (1)
// ═══════════════════════════════════════════════════════

const executeCode = createTool({
  description:
    "Execute Python code for data analysis and transformation. The sandbox has pandas, numpy, datetime pre-imported.",
  args: z.object({
    code: z.string().describe("Python code to execute"),
    summary: z.string().describe("Brief description of what the code does"),
  }),
  handler: async (_ctx, args) => {
    try {
      const { Sandbox } = await import("@e2b/code-interpreter");
      const sbx = await Sandbox.create({
        apiKey: process.env.E2B_API_KEY,
        timeoutMs: 60_000,
      });

      try {
        const execution = await sbx.runCode(args.code, { timeoutMs: 30_000 });
        return {
          stdout: execution.logs.stdout,
          stderr: execution.logs.stderr,
          result: execution.text,
          error: execution.error
            ? {
                name: execution.error.name,
                value: execution.error.value,
                traceback: execution.error.traceback,
              }
            : null,
        };
      } finally {
        await sbx.kill();
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Sandbox creation failed";
      return {
        error: { name: "SandboxError", value: msg, traceback: "" },
      };
    }
  },
});

// ═══════════════════════════════════════════════════════
// Finance (1)
// ═══════════════════════════════════════════════════════

const getStockPrice = createTool({
  description: "Fetch the latest stock price information for a given ticker.",
  args: z.object({
    symbol: z.string().describe("Stock ticker symbol like AAPL, TSLA"),
  }),
  handler: async (_ctx, args) => {
    try {
      const { Sandbox } = await import("@e2b/code-interpreter");
      const sbx = await Sandbox.create({
        apiKey: process.env.E2B_API_KEY,
        timeoutMs: 30_000,
      });

      try {
        const code = `
import yfinance as yf
import json
ticker = yf.Ticker("${args.symbol.replace(/"/g, "")}")
info = ticker.info
result = {
    "symbol": "${args.symbol.replace(/"/g, "")}",
    "price": info.get("currentPrice") or info.get("regularMarketPrice"),
    "previousClose": info.get("previousClose"),
    "open": info.get("open"),
    "dayHigh": info.get("dayHigh"),
    "dayLow": info.get("dayLow"),
    "volume": info.get("volume"),
    "marketCap": info.get("marketCap"),
    "name": info.get("shortName"),
}
print(json.dumps(result))
`;
        const execution = await sbx.runCode(code);
        if (execution.error) {
          return { error: execution.error.value };
        }
        const stdout = execution.logs.stdout.join("");
        return JSON.parse(stdout) as Record<string, string | number | null>;
      } finally {
        await sbx.kill();
      }
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Stock price fetch failed";
      return { error: msg };
    }
  },
});

// ═══════════════════════════════════════════════════════
// Utility (1)
// ═══════════════════════════════════════════════════════

const echoTool = createTool({
  description: "Return a message.",
  args: z.object({
    message: z.string(),
  }),
  handler: async (_ctx, args) => {
    return { message: args.message };
  },
});

// ═══════════════════════════════════════════════════════
// Helper types & functions
// ═══════════════════════════════════════════════════════

interface CellData {
  v?: string | number | boolean;
  f?: string;
  t?: string;
  s?: Record<string, unknown>;
}

type CellReadValue = { value?: string | number | boolean; formula?: string; type?: string } | null;

interface RangeReadResult {
  cells: CellReadValue[][];
  rowCount: number;
  colCount: number;
}

interface SheetData {
  id: string;
  name: string;
  tabColor?: string;
  rowCount?: number;
  columnCount?: number;
  cellData?: Record<string, Record<string, CellData>>;
  dataValidations?: unknown[];
  mergeData?: unknown[];
}

interface WorkbookSnapshot {
  id: string;
  locale?: string;
  name?: string;
  sheetOrder: string[];
  sheets: Record<string, SheetData>;
}

function findSheet(snapshot: WorkbookSnapshot, name: string): SheetData | undefined {
  return Object.values(snapshot.sheets).find((s) => s.name === name);
}

function parseRange(range: string): {
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

function parseCellRef(ref: string): { row: number; col: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error(`Invalid cell reference: ${ref}`);
  const colStr = match[1].toUpperCase();
  const row = parseInt(match[2]) - 1; // 0-indexed
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  col -= 1; // 0-indexed
  return { row, col };
}

function getUsedRange(cellData: Record<string, Record<string, CellData>> | undefined): string | null {
  if (!cellData) return null;

  let minRow = Infinity;
  let maxRow = -1;
  let minCol = Infinity;
  let maxCol = -1;

  for (const [rowStr, rowData] of Object.entries(cellData)) {
    const row = parseInt(rowStr);
    for (const colStr of Object.keys(rowData)) {
      const col = parseInt(colStr);
      minRow = Math.min(minRow, row);
      maxRow = Math.max(maxRow, row);
      minCol = Math.min(minCol, col);
      maxCol = Math.max(maxCol, col);
    }
  }

  if (maxRow === -1) return null;
  return `${colToLetter(minCol)}${minRow + 1}:${colToLetter(maxCol)}${maxRow + 1}`;
}

function colToLetter(col: number): string {
  let letter = "";
  let c = col + 1; // 1-indexed
  while (c > 0) {
    const rem = (c - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    c = Math.floor((c - 1) / 26);
  }
  return letter;
}

// ═══════════════════════════════════════════════════════
// Agent definition
// ═══════════════════════════════════════════════════════

const openrouter = createOpenAI({
  baseURL: "https://openrouter.ai/api/v1",
  apiKey: process.env.OPENROUTER_API_KEY,
});

export const sheetsAgent: Agent<{ workbookId: Id<"workbooks"> }> = new Agent<{ workbookId: Id<"workbooks"> }>(
  components.agent,
  {
    name: "Ramp Sheets Agent",
    languageModel: openrouter.chat("openai/gpt-4o"),
    instructions: SYSTEM_PROMPT,
    tools: {
      // Task Management (4)
      read_tasks: readTasks,
      add_tasks: addTasks,
      complete_task: completeTask,
      delete_tasks: deleteTasks,

      // Spreadsheet Reading (7)
      read_ranges: readRanges,
      get_workbook_metadata: getWorkbookMetadata,
      get_data_validation_info: getDataValidationInfo,
      get_merged_cells: getMergedCells,
      validate_excel_range: validateExcelRange,
      validate_formula_syntax: validateFormulaSyntax,
      grep: grepTool,

      // Spreadsheet Writing (2)
      write_data_to_excel: writeDataToExcel,
      write_df_to_sheet: writeDfToSheet,

      // Formatting (3)
      format_range: formatRange,
      conditional_format_range: conditionalFormatRange,
      format_worksheet: formatWorksheet,

      // Sheet Management (5)
      rename_worksheet: renameWorksheet,
      create_worksheet: createWorksheet,
      delete_worksheet: deleteWorksheet,
      copy_worksheet: copyWorksheet,
      create_table: createTable,

      // Cell Operations (9)
      apply_formula: applyFormula,
      drag_formula: dragFormula,
      copy_range: copyRange,
      delete_range: deleteRange,
      clear_cells: clearCells,
      merge_cells: mergeCells,
      unmerge_cells: unmergeCells,
      set_cell_dropdown: setCellDropdown,
      sort_range: sortRange,

      // Data Operations (1)
      create_pivot_table: createPivotTable,

      // Web & Research (3)
      search_web: searchWeb,
      generate_research_report: generateResearchReport,
      browse_url: browseUrl,

      // Code Execution (1)
      execute_code: executeCode,

      // Finance (1)
      get_stock_price: getStockPrice,

      // Utility (1)
      echo: echoTool,
    },
    maxSteps: 30,
  }
);
