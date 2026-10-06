import { jsonb, pgTable, primaryKey, text, timestamp } from "drizzle-orm/pg-core";

export const appData = pgTable("app_data", {
  owner: text("owner").notNull(),
  key: text("key").notNull(),
  value: jsonb("value").notNull(),
  updatedAt: timestamp("updated_at", { withTimezone: true }).notNull().defaultNow(),
}, table => [
  primaryKey({ columns: [table.owner, table.key] }),
]);

export type AppDataRecord = typeof appData.$inferSelect;
