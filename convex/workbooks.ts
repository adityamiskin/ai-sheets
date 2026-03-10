import { mutation, query } from "./_generated/server";
import { v } from "convex/values";
import { DEFAULT_SNAPSHOT, DEFAULT_WORKBOOK_NAME } from "./defaultWorkbook";

export const create = mutation({
  args: {
    name: v.optional(v.string()),
  },
  handler: async (ctx, args) => {
    const now = Date.now();
    const workbookId = await ctx.db.insert("workbooks", {
      name: args.name ?? DEFAULT_WORKBOOK_NAME,
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
      name: DEFAULT_WORKBOOK_NAME,
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
