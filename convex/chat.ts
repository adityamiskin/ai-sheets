import { mutation, query } from "./_generated/server";
import { v } from "convex/values";
import { internal, components } from "./_generated/api";
import {
  createThread,
  getThreadMetadata,
  listUIMessages,
  saveMessage,
  syncStreams,
  vStreamArgs,
} from "@convex-dev/agent";
import { paginationOptsValidator } from "convex/server";
import { DEFAULT_SNAPSHOT } from "./defaultWorkbook";

// Create a new chat thread for a workbook
export const createChatThread = mutation({
  args: {
    workbookId: v.id("workbooks"),
  },
  handler: async (ctx, args) => {
    const threadId = await createThread(ctx, components.agent, {
      title: "New thread",
    });
    const now = Date.now();
    await ctx.db.insert("chatThreads", {
      workbookId: args.workbookId,
      threadId,
      snapshot: DEFAULT_SNAPSHOT,
      createdAt: now,
      updatedAt: now,
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

    const threadId = await createThread(ctx, components.agent, {
      title: "New thread",
    });
    const now = Date.now();
    await ctx.db.insert("chatThreads", {
      workbookId: args.workbookId,
      threadId,
      snapshot: DEFAULT_SNAPSHOT,
      createdAt: now,
      updatedAt: now,
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
    const thread = await ctx.db
      .query("chatThreads")
      .withIndex("by_thread", (q) => q.eq("threadId", args.threadId))
      .unique();

    if (thread) {
      await ctx.db.patch(thread._id, {
        updatedAt: Date.now(),
      });
    }

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
      .order("desc")
      .first();
    return thread?.threadId ?? null;
  },
});

export const listThreadsForWorkbook = query({
  args: { workbookId: v.id("workbooks") },
  handler: async (ctx, args) => {
    const threads = await ctx.db
      .query("chatThreads")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .order("desc")
      .collect();

    const metadataByThread = await Promise.all(
      threads.map(async (thread) => {
        const metadata = await getThreadMetadata(ctx, components.agent, {
          threadId: thread.threadId,
        }).catch(() => null);
        return [thread.threadId, metadata] as const;
      })
    );

    const metadataMap = new Map(metadataByThread);

    return threads.map((thread) => {
      const metadata = metadataMap.get(thread.threadId);
      return {
        threadId: thread.threadId,
        title: metadata?.title ?? thread.title ?? "New thread",
        summary: metadata?.summary ?? null,
        createdAt: thread.createdAt,
        updatedAt: thread.updatedAt ?? thread.createdAt,
      };
    });
  },
});

export const getThreadState = query({
  args: { threadId: v.string() },
  handler: async (ctx, args) => {
    const thread = await ctx.db
      .query("chatThreads")
      .withIndex("by_thread", (q) => q.eq("threadId", args.threadId))
      .unique();

    if (!thread) return null;

    const metadata = await getThreadMetadata(ctx, components.agent, {
      threadId: args.threadId,
    }).catch(() => null);

    return {
      threadId: thread.threadId,
      title: metadata?.title ?? thread.title ?? "New thread",
      summary: metadata?.summary ?? null,
      snapshot: thread.snapshot ?? DEFAULT_SNAPSHOT,
      createdAt: thread.createdAt,
      updatedAt: thread.updatedAt ?? thread.createdAt,
    };
  },
});

export const updateThreadSnapshot = mutation({
  args: {
    threadId: v.string(),
    snapshot: v.string(),
  },
  handler: async (ctx, args) => {
    const thread = await ctx.db
      .query("chatThreads")
      .withIndex("by_thread", (q) => q.eq("threadId", args.threadId))
      .unique();

    if (!thread) throw new Error(`Thread not found: ${args.threadId}`);

    await ctx.db.patch(thread._id, {
      snapshot: args.snapshot,
      updatedAt: Date.now(),
    });
  },
});
