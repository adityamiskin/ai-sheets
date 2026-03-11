export type SupportedSqlDialect = "postgres" | "mysql" | "sqlserver";

export function detectSqlDialect(connectionString: string): SupportedSqlDialect {
  const protocol = parseConnectionProtocol(connectionString);

  switch (protocol) {
    case "postgres":
    case "postgresql":
      return "postgres";
    case "mysql":
      return "mysql";
    case "sqlserver":
    case "mssql":
      return "sqlserver";
    default:
      throw new Error(
        `Unsupported connection string protocol '${protocol}'. Supported protocols: postgres, postgresql, mysql, sqlserver, mssql.`,
      );
  }
}

export function maskConnectionString(connectionString: string): string {
  return connectionString.replace(
    /^([a-z0-9+]+:\/\/)([^:@\/?#]+)(?::([^@\/?#]*))?@/i,
    (_match, prefix: string, username: string, password?: string) =>
      password !== undefined
        ? `${prefix}${username}:••••••@`
        : `${prefix}${username}@`,
  );
}

function parseConnectionProtocol(connectionString: string) {
  try {
    const url = new URL(connectionString);
    return url.protocol.replace(/:$/, "").toLowerCase();
  } catch {
    const match = connectionString.match(/^([a-z0-9+]+):\/\//i);
    if (!match) {
      throw new Error("Connection string must be a valid URL.");
    }
    return match[1].toLowerCase();
  }
}
