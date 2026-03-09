import { useQuery, useMutation } from "convex/react";
import { api } from "../../convex/_generated/api";
import { useCallback, useRef, useState } from "react";
import type { Id } from "../../convex/_generated/dataModel";

/**
 * Manages workbook CRUD and snapshot persistence with Convex.
 */
export function useWorkbook() {
  const [workbookId, setWorkbookId] = useState<Id<"workbooks"> | null>(null);
  const getOrCreateMutation = useMutation(api.workbooks.getOrCreate);
  const updateSnapshotMutation = useMutation(api.workbooks.updateSnapshot);
  const workbook = useQuery(
    api.workbooks.get,
    workbookId ? { id: workbookId } : "skip"
  );

  const initialized = useRef(false);

  const initialize = useCallback(async () => {
    if (initialized.current) return;
    initialized.current = true;
    const wb = await getOrCreateMutation({});
    setWorkbookId(wb._id);
    return wb;
  }, [getOrCreateMutation]);

  const saveSnapshot = useCallback(
    async (snapshot: string) => {
      if (!workbookId) return;
      await updateSnapshotMutation({ id: workbookId, snapshot });
    },
    [workbookId, updateSnapshotMutation]
  );

  return {
    workbookId,
    workbook,
    initialize,
    saveSnapshot,
  };
}
