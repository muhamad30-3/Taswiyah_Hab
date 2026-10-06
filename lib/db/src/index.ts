import { drizzle } from "drizzle-orm/node-postgres";
import pg from "pg";
import * as schema from "./schema";

const { Pool } = pg;

if (!process.env.DATABASE_URL) {
  throw new Error(
    "DATABASE_URL must be set. Did you forget to provision a database?",
  );
}

export const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  connectionTimeoutMillis: 8_000,
  idleTimeoutMillis: 30_000,
  keepAlive: true,
  keepAliveInitialDelayMillis: 10_000,
  query_timeout: 60_000,
  statement_timeout: 60_000,
});

pool.on("error", error => {
  const errorCode = "code" in error && typeof error.code === "string" ? error.code : error.name;
  console.error("Unexpected idle PostgreSQL connection error:", errorCode);
});

export const db = drizzle(pool, { schema });

export * from "./schema";
