import { internalQuery, mutation, query } from "./_generated/server";
import { v } from "convex/values";
import type { Id } from "./_generated/dataModel";
import {
  detectSqlDialect,
  maskConnectionString,
  type SupportedSqlDialect,
} from "./sqlDataShared";

export const listForWorkbook = query({
  args: {
    workbookId: v.id("workbooks"),
  },
  handler: async (ctx, args) => {
    const sources = await ctx.db
      .query("dataSources")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .order("desc")
      .collect();

    return sources.map((source) => ({
      _id: source._id,
      workbookId: source.workbookId,
      name: source.name,
      dialect: source.dialect,
      maskedConnectionString: maskConnectionString(source.connectionString),
      createdAt: source.createdAt,
      updatedAt: source.updatedAt,
    }));
  },
});

export const create = mutation({
  args: {
    workbookId: v.id("workbooks"),
    name: v.string(),
    connectionString: v.string(),
  },
  handler: async (ctx, args) => {
    const dialect = detectSqlDialect(args.connectionString);
    const now = Date.now();

    const id = await ctx.db.insert("dataSources", {
      workbookId: args.workbookId,
      name: args.name.trim(),
      connectionString: args.connectionString.trim(),
      dialect,
      createdAt: now,
      updatedAt: now,
    });

    return id;
  },
});

export const remove = mutation({
  args: {
    id: v.id("dataSources"),
  },
  handler: async (ctx, args) => {
    await ctx.db.delete(args.id);
  },
});

export const getMaskedById = query({
  args: {
    id: v.id("dataSources"),
  },
  handler: async (ctx, args) => {
    const source = await ctx.db.get(args.id);
    if (!source) {
      return null;
    }

    return {
      _id: source._id,
      workbookId: source.workbookId,
      name: source.name,
      dialect: source.dialect,
      maskedConnectionString: maskConnectionString(source.connectionString),
      createdAt: source.createdAt,
      updatedAt: source.updatedAt,
    };
  },
});

export const internalListForWorkbook = internalQuery({
  args: {
    workbookId: v.id("workbooks"),
  },
  handler: async (ctx, args) => {
    return await ctx.db
      .query("dataSources")
      .withIndex("by_workbook", (q) => q.eq("workbookId", args.workbookId))
      .order("desc")
      .collect();
  },
});

export const internalGetById = internalQuery({
  args: {
    id: v.id("dataSources"),
  },
  handler: async (ctx, args) => {
    return await ctx.db.get(args.id);
  },
});

export type StoredDataSource = {
  _id: Id<"dataSources">;
  workbookId: Id<"workbooks">;
  name: string;
  dialect: SupportedSqlDialect;
  connectionString: string;
  createdAt: number;
  updatedAt: number;
};
