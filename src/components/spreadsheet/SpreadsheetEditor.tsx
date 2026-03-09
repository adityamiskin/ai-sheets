import { useEffect, useRef, useCallback } from "react";
import { createUniver, LocaleType } from "@univerjs/presets";
import { UniverSheetsCorePreset } from "@univerjs/preset-sheets-core";
import sheetsCoreEnUS from "@univerjs/preset-sheets-core/locales/en-US";
import type { FUniver } from "@univerjs/presets";
import type { IWorkbookData } from "@univerjs/core";

import "@univerjs/preset-sheets-core/lib/index.css";

interface SpreadsheetEditorProps {
  /** Initial snapshot to load (JSON string from Convex). */
  snapshot?: string;
  /** Called when the workbook changes (debounced save). */
  onSnapshotChange?: (snapshot: string) => void;
  /** Ref callback so parent can access the Facade API. */
  onApiReady?: (api: FUniver) => void;
}

export function SpreadsheetEditor({
  snapshot,
  onSnapshotChange,
  onApiReady,
}: SpreadsheetEditorProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const apiRef = useRef<FUniver | null>(null);
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const initializedRef = useRef(false);

  // Debounced save
  const scheduleSave = useCallback(() => {
    if (!onSnapshotChange || !apiRef.current) return;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      const workbook = apiRef.current?.getActiveWorkbook();
      if (!workbook) return;
      const data = workbook.save();
      onSnapshotChange(JSON.stringify(data));
    }, 2000);
  }, [onSnapshotChange]);

  useEffect(() => {
    if (!containerRef.current || initializedRef.current) return;
    initializedRef.current = true;

    // Parse snapshot if provided
    let workbookData: Partial<IWorkbookData> | undefined;
    if (snapshot) {
      try {
        workbookData = JSON.parse(snapshot) as Partial<IWorkbookData>;
      } catch {
        console.warn("Failed to parse workbook snapshot, using default");
      }
    }

    const { univerAPI } = createUniver({
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: sheetsCoreEnUS,
      },
      presets: [
        UniverSheetsCorePreset({
          container: containerRef.current,
        }),
      ],
    });

    apiRef.current = univerAPI;

    // Create workbook from snapshot or with default empty sheet
    if (workbookData) {
      univerAPI.createWorkbook(workbookData);
    }

    // Notify parent the API is ready
    onApiReady?.(univerAPI);

    // Listen for changes to auto-save
    univerAPI.addEvent(univerAPI.Event.CommandExecuted, () => {
      scheduleSave();
    });

    return () => {
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
      apiRef.current = null;
      initializedRef.current = false;
    };
    // Only run on mount
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div
      ref={containerRef}
      style={{ width: "100%", height: "100%" }}
    />
  );
}
