import { mutation, query } from "./_generated/server";
import { v } from "convex/values";

// Default empty workbook snapshot for Univer
const DEFAULT_SNAPSHOT = JSON.stringify({
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

export const create = mutation({
  args: {
    name: v.optional(v.string()),
  },
  handler: async (ctx, args) => {
    const now = Date.now();
    const workbookId = await ctx.db.insert("workbooks", {
      name: args.name ?? "Book1.xlsx",
      snapshot: DEFAULT_SNAPSHOT,
      createdAt: now,
      updatedAt: now,
    });
    return workbookId;
  },
});

export const get = query({
  args: { id: v.id("workbooks") },
  handler: async (ctx, args) => {
    return await ctx.db.get(args.id);
  },
});

export const getOrCreate = mutation({
  args: {},
  handler: async (ctx) => {
    // Get the most recent workbook, or create one
    const existing = await ctx.db.query("workbooks").order("desc").first();
    if (existing) return existing;

    const now = Date.now();
    const id = await ctx.db.insert("workbooks", {
      name: "Book1.xlsx",
      snapshot: DEFAULT_SNAPSHOT,
      createdAt: now,
      updatedAt: now,
    });
    return (await ctx.db.get(id))!;
  },
});

export const updateSnapshot = mutation({
  args: {
    id: v.id("workbooks"),
    snapshot: v.string(),
  },
  handler: async (ctx, args) => {
    await ctx.db.patch(args.id, {
      snapshot: args.snapshot,
      updatedAt: Date.now(),
    });
  },
});

export const list = query({
  args: {},
  handler: async (ctx) => {
    return await ctx.db.query("workbooks").order("desc").collect();
  },
});
