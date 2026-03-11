"use node";

import mysql from "mysql2/promise";
import mssql from "mssql";
import postgres from "postgres";
import { type SupportedSqlDialect } from "./sqlDataShared";

export interface SqlRow {
  [key: string]: JsonValue;
}

export type JsonPrimitive = string | number | boolean | null;
export type JsonValue = JsonPrimitive | JsonValue[] | { [key: string]: JsonValue };

export interface SqlQueryResult {
  columns: string[];
  rows: SqlRow[];
}

export interface SqlSource {
  name: string;
  dialect: SupportedSqlDialect;
  connectionString: string;
}

export function assertReadOnlySql(sql: string) {
  const normalized = sql
    .replace(/--.*$/gm, " ")
    .replace(/\/\*[\s\S]*?\*\//g, " ")
    .trim()
    .toLowerCase();

  if (!normalized) {
    throw new Error("SQL query is empty.");
  }

  if (normalized.includes(";")) {
    const stripped = normalized.replace(/;\s*$/, "");
    if (stripped.includes(";")) {
      throw new Error("Only a single SQL statement is allowed.");
    }
  }

  if (
    !(
      normalized.startsWith("select") ||
      normalized.startsWith("with") ||
      normalized.startsWith("show") ||
      normalized.startsWith("describe") ||
      normalized.startsWith("desc") ||
      normalized.startsWith("explain")
    )
  ) {
    throw new Error("Only read-only SQL statements are allowed.");
  }

  if (
    /\b(insert|update|delete|drop|alter|truncate|grant|revoke|merge|create|replace|upsert|call|copy)\b/.test(
      normalized,
    )
  ) {
    throw new Error("Mutating SQL keywords are not allowed.");
  }
}

export async function runSqlQuery(
  source: SqlSource,
  sql: string,
): Promise<SqlQueryResult> {
  switch (source.dialect) {
    case "postgres":
      return await runPostgresQuery(source.connectionString, sql);
    case "mysql":
      return await runMysqlQuery(source.connectionString, sql);
    case "sqlserver":
      return await runSqlServerQuery(source.connectionString, sql);
  }
}

export async function discoverDatabaseObjects(
  source: SqlSource,
  options: {
    schema?: string;
    namePattern?: string;
    limit?: number;
  } = {},
): Promise<Array<{ schema: string; name: string; type: "table" | "view" }>> {
  const limit = normalizeLimit(options.limit ?? 50, 1, 200);
  const schema = sanitizeSchemaName(
    options.schema ?? inferDefaultSchema(source),
  );
  const namePattern = sanitizeLikePattern(options.namePattern ?? "%");

  switch (source.dialect) {
    case "postgres":
      return await queryAsObjects(source, `
        select table_schema as schema, table_name as name,
          case
            when table_type = 'VIEW' then 'view'
            else 'table'
          end as type
        from information_schema.tables
        where table_schema = '${schema}'
          and table_name ilike '${namePattern}'
        order by table_name
        limit ${limit}
      `);
    case "mysql":
      return await queryAsObjects(source, `
        select table_schema as schema, table_name as name,
          case
            when table_type = 'VIEW' then 'view'
            else 'table'
          end as type
        from information_schema.tables
        where table_schema = '${schema}'
          and table_name like '${namePattern}'
        order by table_name
        limit ${limit}
      `);
    case "sqlserver":
      return await queryAsObjects(source, `
        select top (${limit})
          table_schema as schema,
          table_name as name,
          case
            when table_type = 'VIEW' then 'view'
            else 'table'
          end as type
        from information_schema.tables
        where table_schema = '${schema}'
          and table_name like '${namePattern}'
        order by table_name
      `);
  }
}

export async function inspectRelations(
  source: SqlSource,
  relationNames: string[],
  schema?: string,
): Promise<
  Array<{
    schema: string;
    name: string;
    type: "table" | "view" | "unknown";
    columns: Array<{
      name: string;
      dataType: string;
      nullable: boolean;
      defaultValue: JsonValue;
    }>;
  }>
> {
  const fallbackSchema = sanitizeSchemaName(schema ?? inferDefaultSchema(source));

  const results = [];

  for (const rawName of relationNames) {
    const parsed = parseRelationName(rawName, fallbackSchema);
    const rows = await queryAsObjects(source, relationColumnsQuery(source.dialect, parsed.schema, parsed.name));

    const relationType = rows[0]?.table_type;
    const type: "table" | "view" | "unknown" =
      relationType === "VIEW"
        ? "view"
        : relationType === "BASE TABLE"
          ? "table"
          : "unknown";

    results.push({
      schema: parsed.schema,
      name: parsed.name,
      type,
      columns: rows.map((row) => ({
        name: String(row.column_name ?? ""),
        dataType: String(row.data_type ?? ""),
        nullable: String(row.is_nullable ?? "").toUpperCase() === "YES",
        defaultValue: toJsonValue(row.column_default ?? null),
      })),
    });
  }

  return results;
}

function relationColumnsQuery(
  dialect: SupportedSqlDialect,
  schema: string,
  tableName: string,
) {
  const escapedSchema = sanitizeSchemaName(schema);
  const escapedTable = sanitizeIdentifier(tableName);

  switch (dialect) {
    case "postgres":
    case "mysql":
    case "sqlserver":
      return `
        select
          c.table_schema,
          c.table_name,
          c.column_name,
          c.data_type,
          c.is_nullable,
          c.column_default,
          t.table_type
        from information_schema.columns c
        left join information_schema.tables t
          on t.table_schema = c.table_schema
         and t.table_name = c.table_name
        where c.table_schema = '${escapedSchema}'
          and c.table_name = '${escapedTable}'
        order by c.ordinal_position
      `;
  }
}

async function queryAsObjects<T extends Record<string, unknown>>(
  source: SqlSource,
  sql: string,
): Promise<T[]> {
  const result = await runSqlQuery(source, sql);
  return result.rows as T[];
}

async function runPostgresQuery(
  connectionString: string,
  sql: string,
): Promise<SqlQueryResult> {
  const client = postgres(connectionString, {
    max: 1,
    idle_timeout: 5,
    connect_timeout: 10,
    prepare: false,
  });

  try {
    const rows = await client.unsafe<Record<string, unknown>[]>(sql);
    const normalizedRows = rows.map((row) => normalizeRow(row));
    return {
      columns: normalizedRows.length > 0 ? Object.keys(normalizedRows[0]) : [],
      rows: normalizedRows,
    };
  } finally {
    await client.end({ timeout: 5 });
  }
}

async function runMysqlQuery(
  connectionString: string,
  sql: string,
): Promise<SqlQueryResult> {
  const connection = await mysql.createConnection(connectionString);

  try {
    const [rows] = await connection.query(sql);
    const normalizedRows = Array.isArray(rows)
      ? rows.map((row) => normalizeRow(row as Record<string, unknown>))
      : [];

    return {
      columns: normalizedRows.length > 0 ? Object.keys(normalizedRows[0]) : [],
      rows: normalizedRows,
    };
  } finally {
    await connection.end();
  }
}

async function runSqlServerQuery(
  connectionString: string,
  sql: string,
): Promise<SqlQueryResult> {
  const pool = new mssql.ConnectionPool(
    parseSqlServerConnectionString(connectionString),
  );
  await pool.connect();

  try {
    const result = await pool.request().query(sql);
    const normalizedRows = (result.recordset ?? []).map(
      (row: Record<string, unknown>) => normalizeRow(row),
    );
    return {
      columns: normalizedRows.length > 0 ? Object.keys(normalizedRows[0]) : [],
      rows: normalizedRows,
    };
  } finally {
    await pool.close();
  }
}

function parseSqlServerConnectionString(
  connectionString: string,
): mssql.config {
  let url: URL;

  try {
    url = new URL(connectionString);
  } catch {
    throw new Error(
      "SQL Server connection string must be a valid sqlserver:// or mssql:// URL.",
    );
  }

  const server = url.hostname;
  if (!server) {
    throw new Error("SQL Server connection string is missing a server hostname.");
  }

  const rawDatabase = url.pathname.replace(/^\/+/, "").trim();
  const database = rawDatabase || undefined;
  const port = url.port ? Number(url.port) : undefined;
  const username = url.username ? decodeURIComponent(url.username) : undefined;
  const password = url.password ? decodeURIComponent(url.password) : undefined;
  const instanceName = url.searchParams.get("instance")?.trim() || undefined;

  return {
    server,
    port,
    user: username,
    password,
    database,
    options: {
      encrypt: parseBooleanParam(url.searchParams.get("encrypt"), true),
      trustServerCertificate: parseBooleanParam(
        url.searchParams.get("trustServerCertificate"),
        true,
      ),
      ...(instanceName ? { instanceName } : {}),
    },
    pool: {
      max: 1,
      min: 0,
      idleTimeoutMillis: 5000,
    },
  };
}

function parseBooleanParam(value: string | null, fallback: boolean) {
  if (value === null) {
    return fallback;
  }

  switch (value.trim().toLowerCase()) {
    case "true":
    case "1":
    case "yes":
      return true;
    case "false":
    case "0":
    case "no":
      return false;
    default:
      return fallback;
  }
}

function normalizeRow(row: Record<string, unknown>): SqlRow {
  return Object.fromEntries(
    Object.entries(row).map(([key, value]) => [key, toJsonValue(value)]),
  );
}

function toJsonValue(value: unknown): JsonValue {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (typeof value === "bigint") {
    return value.toString();
  }
  if (Buffer.isBuffer(value)) {
    return value.toString("base64");
  }
  if (Array.isArray(value)) {
    return value.map((item) => toJsonValue(item));
  }
  if (typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>).map(([key, item]) => [
        key,
        toJsonValue(item),
      ]),
    );
  }
  return String(value);
}

function sanitizeLikePattern(value: string) {
  const sanitized = value.replace(/[^a-zA-Z0-9_%]/g, "");
  return sanitized || "%";
}

function sanitizeSchemaName(value: string) {
  const sanitized = value.replace(/[^a-zA-Z0-9_]/g, "");
  if (!sanitized) {
    throw new Error("Schema name is invalid.");
  }
  return sanitized;
}

function sanitizeIdentifier(value: string) {
  const sanitized = value.replace(/[^a-zA-Z0-9_]/g, "");
  if (!sanitized) {
    throw new Error("Identifier is invalid.");
  }
  return sanitized;
}

function parseRelationName(value: string, fallbackSchema: string) {
  const parts = value.split(".");
  if (parts.length === 1) {
    return {
      schema: fallbackSchema,
      name: sanitizeIdentifier(parts[0]),
    };
  }

  if (parts.length === 2) {
    return {
      schema: sanitizeSchemaName(parts[0]),
      name: sanitizeIdentifier(parts[1]),
    };
  }

  throw new Error(`Invalid relation name '${value}'.`);
}

function inferDefaultSchema(source: SqlSource) {
  if (source.dialect === "mysql") {
    try {
      const url = new URL(source.connectionString);
      const databaseName = url.pathname.replace(/^\/+/, "").split("/")[0];
      return databaseName || "mysql";
    } catch {
      return "mysql";
    }
  }

  switch (source.dialect) {
    case "postgres":
      return "public";
    case "sqlserver":
      return "dbo";
  }
}

function normalizeLimit(value: number, min: number, max: number) {
  return Math.max(min, Math.min(max, Math.floor(value)));
}
