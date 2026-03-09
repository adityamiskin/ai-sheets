"use node";

import { internalAction } from "./_generated/server";
import { v } from "convex/values";
import { sheetsAgent } from "./_agent";

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

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const agent = sheetsAgent as any;
    const result = await agent.streamText(
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

    await result.consumeStream();
  },
});
