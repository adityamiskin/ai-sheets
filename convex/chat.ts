import { query, mutation } from "./_generated/server";
import { v } from "convex/values";
import { internal, components } from "./_generated/api";
import {
  createThread,
  listUIMessages,
  saveMessage,
  syncStreams,
  vStreamArgs,
} from "@convex-dev/agent";
import { paginationOptsValidator } from "convex/server";

// Create a new chat thread for a workbook
export const createChatThread = mutation({
  args: {
    workbookId: v.id("workbooks"),
  },
  handler: async (ctx, args) => {
    const threadId = await createThread(ctx, components.agent, {});
    await ctx.db.insert("chatThreads", {
      workbookId: args.workbookId,
      threadId,
      createdAt: Date.now(),
    });
    return threadId;
  },
});

// Get or create thread for a workbook
export const getOrCreateThread = mutation({
  args: {
    workbookId: v.id("workbooks"),
  },
  handler: async (ctx, args) => {
    const existing = await ctx.db
      .query("chatThreads")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .first();

    if (existing) return existing.threadId;

    const threadId = await createThread(ctx, components.agent, {});
    await ctx.db.insert("chatThreads", {
      workbookId: args.workbookId,
      threadId,
      createdAt: Date.now(),
    });
    return threadId;
  },
});

// Send a message (mutation -> schedule action)
export const sendMessage = mutation({
  args: {
    threadId: v.string(),
    workbookId: v.id("workbooks"),
    prompt: v.string(),
  },
  handler: async (ctx, args) => {
    // Save the user message using the standalone function (no Node.js deps)
    const { messageId } = await saveMessage(ctx, components.agent, {
      threadId: args.threadId,
      prompt: args.prompt,
    });

    // Schedule the AI response action
    await ctx.scheduler.runAfter(0, internal.chatActions.generateResponse, {
      threadId: args.threadId,
      workbookId: args.workbookId,
      promptMessageId: messageId,
    });

    return messageId;
  },
});

// List messages with streaming support
export const listMessages = query({
  args: {
    threadId: v.string(),
    paginationOpts: paginationOptsValidator,
    streamArgs: vStreamArgs,
  },
  handler: async (ctx, args) => {
    const paginated = await listUIMessages(ctx, components.agent, args);
    const streams = await syncStreams(ctx, components.agent, args);
    return { ...paginated, streams };
  },
});

// Get thread for workbook
export const getThreadForWorkbook = query({
  args: { workbookId: v.id("workbooks") },
  handler: async (ctx, args) => {
    const thread = await ctx.db
      .query("chatThreads")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .first();
    return thread?.threadId ?? null;
  },
});
