import { defineSchema, defineTable } from "convex/server";
import { v } from "convex/values";

export default defineSchema({
  // Workbooks - stores the entire spreadsheet state
  workbooks: defineTable({
    name: v.string(),
    // Serialized Univer workbook snapshot (IWorkbookData JSON)
    snapshot: v.string(),
    createdAt: v.number(),
    updatedAt: v.number(),
  }),

  // Operations queue: AI agent writes ops -> frontend applies to Univer
  operations: defineTable({
    workbookId: v.id("workbooks"),
    threadId: v.string(),
    type: v.string(), // tool/operation name
    params: v.string(), // JSON-encoded operation parameters
    status: v.union(
      v.literal("pending"),
      v.literal("applied"),
      v.literal("failed")
    ),
    result: v.optional(v.string()), // JSON-encoded result for read operations
    error: v.optional(v.string()),
    explanation: v.optional(v.string()), // human-readable description
    createdAt: v.number(),
  })
    .index("by_workbook_status", ["workbookId", "status"])
    .index("by_workbook_created", ["workbookId", "createdAt"])
    .index("by_thread", ["threadId", "createdAt"]),

  // Task management for AI
  tasks: defineTable({
    workbookId: v.id("workbooks"),
    threadId: v.optional(v.string()),
    content: v.string(),
    status: v.union(
      v.literal("pending"),
      v.literal("in_progress"),
      v.literal("completed")
    ),
    order: v.number(),
  }).index("by_workbook", ["workbookId", "order"]),

  // Chat threads metadata (agent component manages actual messages)
  chatThreads: defineTable({
    workbookId: v.id("workbooks"),
    threadId: v.string(),
    title: v.optional(v.string()),
    snapshot: v.optional(v.string()),
    createdAt: v.number(),
    updatedAt: v.optional(v.number()),
  })
    .index("by_workbook", ["workbookId"])
    .index("by_thread", ["threadId"]),

  // Store code execution results/state
  codeExecutions: defineTable({
    workbookId: v.id("workbooks"),
    threadId: v.string(),
    code: v.string(),
    stdout: v.optional(v.string()),
    stderr: v.optional(v.string()),
    result: v.optional(v.string()),
    error: v.optional(v.string()),
    status: v.union(
      v.literal("running"),
      v.literal("completed"),
      v.literal("failed")
    ),
    createdAt: v.number(),
  }).index("by_workbook", ["workbookId"]),

  dataSources: defineTable({
    workbookId: v.id("workbooks"),
    name: v.string(),
    dialect: v.union(
      v.literal("postgres"),
      v.literal("mysql"),
      v.literal("sqlserver")
    ),
    connectionString: v.string(),
    createdAt: v.number(),
    updatedAt: v.number(),
  })
    .index("by_workbook", ["workbookId"])
    .index("by_workbook_name", ["workbookId", "name"]),
});
