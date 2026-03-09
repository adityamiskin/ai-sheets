import { mutation, query } from "./_generated/server";
import { v } from "convex/values";

export const create = mutation({
  args: {
    workbookId: v.id("workbooks"),
    threadId: v.string(),
    type: v.string(),
    params: v.string(),
    explanation: v.optional(v.string()),
  },
  handler: async (ctx, args) => {
    const id = await ctx.db.insert("operations", {
      workbookId: args.workbookId,
      threadId: args.threadId,
      type: args.type,
      params: args.params,
      status: "pending",
      explanation: args.explanation,
      createdAt: Date.now(),
    });
    return id;
  },
});

export const createWithResult = mutation({
  args: {
    workbookId: v.id("workbooks"),
    threadId: v.string(),
    type: v.string(),
    params: v.string(),
    result: v.optional(v.string()),
    explanation: v.optional(v.string()),
  },
  handler: async (ctx, args) => {
    const id = await ctx.db.insert("operations", {
      workbookId: args.workbookId,
      threadId: args.threadId,
      type: args.type,
      params: args.params,
      status: "applied", // read-only ops are immediately "applied"
      result: args.result,
      explanation: args.explanation,
      createdAt: Date.now(),
    });
    return id;
  },
});

export const getPending = query({
  args: { workbookId: v.id("workbooks") },
  handler: async (ctx, args) => {
    return await ctx.db
      .query("operations")
      .withIndex("by_workbook_status", (q) =>
        q.eq("workbookId", args.workbookId).eq("status", "pending")
      )
      .collect();
  },
});

export const getByThread = query({
  args: { threadId: v.string() },
  handler: async (ctx, args) => {
    return await ctx.db
      .query("operations")
      .withIndex("by_thread", (q) => q.eq("threadId", args.threadId))
      .collect();
  },
});

export const markApplied = mutation({
  args: { id: v.id("operations") },
  handler: async (ctx, args) => {
    await ctx.db.patch(args.id, { status: "applied" });
  },
});

export const markFailed = mutation({
  args: {
    id: v.id("operations"),
    error: v.string(),
  },
  handler: async (ctx, args) => {
    await ctx.db.patch(args.id, {
      status: "failed",
      error: args.error,
    });
  },
});

export const getRecent = query({
  args: {
    workbookId: v.id("workbooks"),
    limit: v.optional(v.number()),
  },
  handler: async (ctx, args) => {
    return await ctx.db
      .query("operations")
      .withIndex("by_workbook_created", (q) =>
        q.eq("workbookId", args.workbookId)
      )
      .order("desc")
      .take(args.limit ?? 50);
  },
});
