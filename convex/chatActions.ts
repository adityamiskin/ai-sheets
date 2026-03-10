"use node";

import { internalAction } from "./_generated/server";
import { v } from "convex/values";
import { sheetsAgent } from "./_agent";
import { z } from "zod";

// AI response generation (internal action — needs Node.js for OpenAI + E2B)
export const generateResponse = internalAction({
  args: {
    threadId: v.string(),
    workbookId: v.id("workbooks"),
    promptMessageId: v.string(),
  },
  handler: async (ctx, args) => {
    // Spread workbookId onto ctx so tool handlers can access it via wbId()
    const ctxWithWorkbook = { ...ctx, workbookId: args.workbookId };

    await sheetsAgent.streamText<undefined>(
      ctxWithWorkbook,
      {
        threadId: args.threadId,
      },
      {
        promptMessageId: args.promptMessageId,
        system: `Current workbook ID: ${args.workbookId}`,
      },
      {
        contextOptions: {
          recentMessages: 20,
        },
        saveStreamDeltas: {
          chunking: "word",
          throttleMs: 100,
        },
      }
    );

    const { thread } = await sheetsAgent.continueThread(ctxWithWorkbook, {
      threadId: args.threadId,
    });
    const metadata = await thread.getMetadata();
    const existingTitle = metadata.title?.trim();

    if (!existingTitle || existingTitle === "New thread") {
      const result = await thread.generateObject(
        {
          prompt:
            "Generate a short title for this spreadsheet thread. Focus on the user's actual task. Use 3 to 8 words, plain text, no quotes.",
          schema: z.object({
            title: z.string().min(3).max(80),
            summary: z.string().min(8).max(240),
          }),
          schemaDescription:
            "A concise spreadsheet-task title and short summary for the thread.",
        },
        {
          storageOptions: { saveMessages: "none" },
        }
      );

      await thread.updateMetadata({
        title: result.object.title.trim(),
        summary: result.object.summary.trim(),
      });
    }
  },
});

export const deleteAgentThread = internalAction({
  args: {
    threadId: v.string(),
    workbookId: v.id("workbooks"),
  },
  handler: async (ctx, args) => {
    const ctxWithWorkbook = { ...ctx, workbookId: args.workbookId };

    try {
      await sheetsAgent.deleteThreadAsync(ctxWithWorkbook, {
        threadId: args.threadId,
      });
    } catch (error) {
      console.warn(`Failed to delete agent thread ${args.threadId}`, error);
    }
  },
});
