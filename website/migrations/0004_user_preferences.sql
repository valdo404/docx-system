-- Migration 0004: Add preferences column to tenant table (idempotent)
-- Create a new table with the column, copy data, and swap
CREATE TABLE IF NOT EXISTS "tenant_new" (
    "id" TEXT PRIMARY KEY NOT NULL,
    "userId" TEXT NOT NULL UNIQUE,
    "displayName" TEXT,
    "gcsPrefix" TEXT NOT NULL,
    "storageQuotaBytes" INTEGER NOT NULL DEFAULT 104857600,
    "storageUsedBytes" INTEGER NOT NULL DEFAULT 0,
    "createdAt" TEXT NOT NULL,
    "updatedAt" TEXT NOT NULL,
    "preferences" TEXT DEFAULT '{}',
    FOREIGN KEY ("userId") REFERENCES "user"("id") ON DELETE CASCADE
);

INSERT OR IGNORE INTO "tenant_new" ("id", "userId", "displayName", "gcsPrefix", "storageQuotaBytes", "storageUsedBytes", "createdAt", "updatedAt", "preferences")
SELECT "id", "userId", "displayName", "gcsPrefix", "storageQuotaBytes", "storageUsedBytes", "createdAt", "updatedAt", COALESCE("preferences", '{}')
FROM "tenant";

DROP TABLE IF EXISTS "tenant_old";
ALTER TABLE "tenant" RENAME TO "tenant_old";
ALTER TABLE "tenant_new" RENAME TO "tenant";
DROP TABLE IF EXISTS "tenant_old";

CREATE INDEX IF NOT EXISTS "idx_tenant_userId" ON "tenant"("userId");
