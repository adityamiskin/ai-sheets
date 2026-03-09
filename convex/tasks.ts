import { mutation, query } from "./_generated/server";
import { v } from "convex/values";

export const list = query({
  args: { workbookId: v.id("workbooks") },
  handler: async (ctx, args) => {
    return await ctx.db
      .query("tasks")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .collect();
  },
});

export const add = mutation({
  args: {
    workbookId: v.id("workbooks"),
    threadId: v.optional(v.string()),
    tasks: v.array(
      v.object({
        content: v.string(),
      })
    ),
  },
  handler: async (ctx, args) => {
    // Get the max order
    const existing = await ctx.db
      .query("tasks")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .collect();
    let maxOrder = existing.reduce((max, t) => Math.max(max, t.order), 0);

    const ids = [];
    for (const task of args.tasks) {
      maxOrder++;
      const id = await ctx.db.insert("tasks", {
        workbookId: args.workbookId,
        threadId: args.threadId,
        content: task.content,
        status: "pending",
        order: maxOrder,
      });
      ids.push(id);
    }
    return ids;
  },
});

export const complete = mutation({
  args: {
    workbookId: v.id("workbooks"),
    taskNumber: v.number(),
  },
  handler: async (ctx, args) => {
    const tasks = await ctx.db
      .query("tasks")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .collect();

    const task = tasks[args.taskNumber - 1]; // 1-indexed
    if (!task) throw new Error(`Task ${args.taskNumber} not found`);

    await ctx.db.patch(task._id, { status: "completed" });
  },
});

export const deleteAll = mutation({
  args: { workbookId: v.id("workbooks") },
  handler: async (ctx, args) => {
    const tasks = await ctx.db
      .query("tasks")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .collect();

    for (const task of tasks) {
      await ctx.db.delete(task._id);
    }
  },
});

export const deleteSpecific = mutation({
  args: {
    workbookId: v.id("workbooks"),
    taskNumbers: v.array(v.number()),
  },
  handler: async (ctx, args) => {
    const tasks = await ctx.db
      .query("tasks")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .collect();

    for (const num of args.taskNumbers) {
      const task = tasks[num - 1]; // 1-indexed
      if (task) {
        await ctx.db.delete(task._id);
      }
    }
  },
});

export const updateStatus = mutation({
  args: {
    id: v.id("tasks"),
    status: v.union(
      v.literal("pending"),
      v.literal("in_progress"),
      v.literal("completed")
    ),
  },
  handler: async (ctx, args) => {
    await ctx.db.patch(args.id, { status: args.status });
  },
});
