import { action, internalMutation } from "./_generated/server";
import { v } from "convex/values";
import { internal } from "./_generated/api";

export const deleteThreadRecords = internalMutation({
  args: {
    threadId: v.string(),
  },
  handler: async (ctx, args) => {
    const thread = await ctx.db
      .query("chatThreads")
      .withIndex("by_thread", (q) => q.eq("threadId", args.threadId))
      .unique();

    if (!thread) {
      return { nextThreadId: null, workbookId: null };
    }

    const operations = await ctx.db
      .query("operations")
      .withIndex("by_thread", (q) => q.eq("threadId", args.threadId))
      .collect();
    for (const operation of operations) {
      await ctx.db.delete(operation._id);
    }

    const tasks = await ctx.db
      .query("tasks")
      .withIndex("by_workbook", (q) => q.eq("workbookId", thread.workbookId))
      .collect();
    for (const task of tasks) {
      if (task.threadId === args.threadId) {
        await ctx.db.delete(task._id);
      }
    }

    const codeExecutions = await ctx.db
      .query("codeExecutions")
      .withIndex("by_workbook", (q) => q.eq("workbookId", thread.workbookId))
      .collect();
    for (const execution of codeExecutions) {
      if (execution.threadId === args.threadId) {
        await ctx.db.delete(execution._id);
      }
    }

    await ctx.db.delete(thread._id);

    const nextThread = await ctx.db
      .query("chatThreads")
      .withIndex("by_workbook", (q) => q.eq("workbookId", thread.workbookId))
      .order("desc")
      .first();

    return {
      nextThreadId: nextThread?.threadId ?? null,
      workbookId: thread.workbookId,
    };
  },
});

export const deleteThread = action({
  args: {
    threadId: v.string(),
  },
  handler: async (ctx, args): Promise<{ nextThreadId: string | null; workbookId: string | null }> => {
    const result = await ctx.runMutation(internal.threadManagement.deleteThreadRecords, {
      threadId: args.threadId,
    });

    if (result.workbookId) {
      await ctx.runAction(internal.chatActions.deleteAgentThread, {
        threadId: args.threadId,
        workbookId: result.workbookId,
      });
    }

    return result;
  },
});
