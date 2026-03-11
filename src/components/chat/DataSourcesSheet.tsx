import { useMemo, useState } from "react";
import { useMutation, useQuery } from "convex/react";
import { Database, Trash2 } from "lucide-react";
import { api } from "../../../convex/_generated/api";
import type { Id } from "../../../convex/_generated/dataModel";
import {
  Sheet,
  SheetContent,
  SheetDescription,
  SheetHeader,
  SheetTitle,
} from "@/components/ui/sheet";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";

interface DataSourcesSheetProps {
  workbookId: Id<"workbooks"> | null;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export function DataSourcesSheet({
  workbookId,
  open,
  onOpenChange,
}: DataSourcesSheetProps) {
  const sources = useQuery(
    api.dataSources.listForWorkbook,
    workbookId ? { workbookId } : "skip",
  );
  const createSource = useMutation(api.dataSources.create);
  const removeSource = useMutation(api.dataSources.remove);
  const [name, setName] = useState("");
  const [connectionString, setConnectionString] = useState("");
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const isFormValid = useMemo(
    () => Boolean(workbookId && name.trim() && connectionString.trim()),
    [connectionString, name, workbookId],
  );

  const handleCreate = async () => {
    if (!workbookId || !isFormValid) {
      return;
    }

    setIsSaving(true);
    setError(null);

    try {
      await createSource({
        workbookId,
        name: name.trim(),
        connectionString: connectionString.trim(),
      });
      setName("");
      setConnectionString("");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to add data source.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleRemove = async (id: Id<"dataSources">) => {
    try {
      await removeSource({ id });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to remove data source.");
    }
  };

  return (
    <Sheet open={open} onOpenChange={onOpenChange}>
      <SheetContent side="right" className="w-full sm:max-w-xl">
        <SheetHeader className="border-b border-[var(--chat-border)]">
          <SheetTitle>Data Sources</SheetTitle>
          <SheetDescription>
            Add one or more SQL connection strings. The agent can discover schema,
            write SQL, and place results into the sheet.
          </SheetDescription>
        </SheetHeader>

        <div className="flex min-h-0 flex-1 flex-col overflow-hidden">
          <div className="space-y-3 border-b border-[var(--chat-border)] px-4 py-4">
            <div className="space-y-1.5">
              <label className="text-xs font-medium text-[var(--soft-text)]">
                Name
              </label>
              <Input
                value={name}
                onChange={(event) => setName(event.target.value)}
                placeholder="Analytics Warehouse"
              />
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-medium text-[var(--soft-text)]">
                Connection string
              </label>
              <Textarea
                value={connectionString}
                onChange={(event) => setConnectionString(event.target.value)}
                placeholder="postgresql://user:password@host:5432/dbname"
                className="min-h-[120px] font-mono text-xs"
              />
            </div>

            {error ? (
              <div className="text-xs text-red-600">{error}</div>
            ) : (
              <div className="text-xs text-[var(--soft-text)]">
                Supported protocols: `postgresql://`, `mysql://`, `sqlserver://`,
                `mssql://`
              </div>
            )}

            <Button
              type="button"
              variant="outline"
              disabled={!isFormValid || isSaving}
              onClick={handleCreate}
            >
              <Database className="h-3.5 w-3.5" />
              Add source
            </Button>
          </div>

          <div className="min-h-0 flex-1 overflow-y-auto px-4 py-4">
            <div className="mb-3 text-xs font-medium uppercase tracking-[0.18em] text-[var(--soft-text)]">
              Connected sources
            </div>

            <div className="space-y-3">
              {sources === undefined ? (
                <div className="text-sm text-[var(--soft-text)]">Loading...</div>
              ) : sources.length === 0 ? (
                <div className="rounded-xl border border-dashed border-[var(--chat-border)] bg-white px-4 py-5 text-sm text-[var(--soft-text)]">
                  No sources yet. Add one above, then ask the AI to inspect or query it.
                </div>
              ) : (
                sources.map((source) => (
                  <div
                    key={source._id}
                    className="rounded-2xl border border-[var(--chat-border)] bg-white p-4"
                  >
                    <div className="flex items-start justify-between gap-3">
                      <div className="min-w-0">
                        <div className="truncate text-sm font-medium text-[var(--ink)]">
                          {source.name}
                        </div>
                        <div className="mt-1 inline-flex rounded-full bg-[var(--chat-soft)] px-2 py-0.5 text-[11px] uppercase tracking-[0.16em] text-[var(--soft-text)]">
                          {source.dialect}
                        </div>
                      </div>

                      <Button
                        type="button"
                        variant="ghost"
                        size="icon-sm"
                        onClick={() => void handleRemove(source._id)}
                        aria-label={`Remove ${source.name}`}
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </Button>
                    </div>

                    <div className="mt-3 break-all rounded-xl bg-[var(--chat-soft)] px-3 py-2 font-mono text-[11px] text-[var(--soft-text)]">
                      {source.maskedConnectionString}
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      </SheetContent>
    </Sheet>
  );
}
