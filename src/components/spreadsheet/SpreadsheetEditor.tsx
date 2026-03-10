import { useCallback, useEffect, useRef } from "react";
import * as GC from "@mescius/spread-sheets";
import "@mescius/spread-sheets-barcode";
import "@mescius/spread-sheets-charts";
import "@mescius/spread-sheets-designer-resources-en";
import "@mescius/spread-sheets-designer";
import "@mescius/spread-sheets-io";
import "@mescius/spread-sheets-languagepackages";
import "@mescius/spread-sheets-pdf";
import "@mescius/spread-sheets-print";
import "@mescius/spread-sheets-shapes";
import "@mescius/spread-sheets-tablesheet";
import "@mescius/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import "@mescius/spread-sheets-designer/styles/gc.spread.sheets.designer.light.min.css";

export type SpreadsheetApi = GC.Spread.Sheets.Workbook;

interface DesignerTabConfig {
  id: string;
  text?: string;
  [key: string]: unknown;
}

interface DesignerConfig {
  ribbon?: DesignerTabConfig[];
  [key: string]: unknown;
}

interface DesignerInstance {
  activeRibbonTab(ribbonTabId?: string): string;
  destroy(): void;
  getWorkbook(): unknown;
  refresh(): void;
}

interface DesignerStatic {
  DefaultConfig: DesignerConfig;
  Designer: new (
    host: HTMLDivElement | string,
    config?: DesignerConfig,
    spread?: object,
    spreadOptions?: object,
  ) => DesignerInstance;
  LicenseKey: string;
}

interface SpreadsheetEditorProps {
  snapshot?: string;
  onSnapshotChange?: (snapshot: string) => void;
  onApiReady?: (api: SpreadsheetApi) => void;
}

const DESIGNER_TAB_IDS = new Set([
  "home",
  "insert",
  "pageLayout",
  "formulas",
  "data",
  "view",
]);
const MIN_SHEET_COLUMN_COUNT = 200;
const MIN_SHEET_ROW_COUNT = 1000;
const SAVE_DEBOUNCE_MS = 1200;

function parseSnapshot(snapshot?: string): unknown | null {
  if (!snapshot) return null;

  try {
    return JSON.parse(snapshot);
  } catch {
    console.warn("Failed to parse workbook snapshot, falling back to a blank workbook");
    return null;
  }
}

function isLegacyUniverSnapshot(
  value: unknown,
): value is { sheetOrder: string[]; sheets: Record<string, unknown> } {
  return (
    typeof value === "object" &&
    value !== null &&
    "sheetOrder" in value &&
    "sheets" in value
  );
}

function isSpreadSnapshot(value: unknown): value is Record<string, unknown> {
  return (
    typeof value === "object" &&
    value !== null &&
    "sheets" in value &&
    !("sheetOrder" in value)
  );
}

function configureSheet(sheet: GC.Spread.Sheets.Worksheet) {
  if (sheet.getColumnCount() < MIN_SHEET_COLUMN_COUNT) {
    sheet.setColumnCount(MIN_SHEET_COLUMN_COUNT);
  }
  if (sheet.getRowCount() < MIN_SHEET_ROW_COUNT) {
    sheet.setRowCount(MIN_SHEET_ROW_COUNT);
  }

  sheet.options.gridline = {
    color: "#e7ddcf",
    showHorizontalGridline: true,
    showVerticalGridline: true,
  };
}

function configureWorkbook(spread: SpreadsheetApi) {
  spread.options.allowUndo = true;
  spread.options.tabStripPosition = GC.Spread.Sheets.TabStripPosition.bottom;
  spread.options.showHorizontalScrollbar = true;
  spread.options.showVerticalScrollbar = true;
  spread.options.newTabVisible = true;

  for (let index = 0; index < spread.getSheetCount(); index += 1) {
    const sheet = spread.getSheet(index);
    if (sheet) {
      configureSheet(sheet);
    }
  }

  const firstSheet = spread.getSheet(0) ?? spread.getActiveSheet();
  if (!firstSheet) return;

  firstSheet.name(firstSheet.name() || "Sheet1");
}

function getDesignerNamespace(): DesignerStatic {
  return (GC as unknown as { Spread: { Sheets: { Designer: DesignerStatic } } }).Spread.Sheets
    .Designer;
}

function createDesignerConfig(): DesignerConfig {
  const defaultConfig = getDesignerNamespace().DefaultConfig;

  return {
    ...defaultConfig,
    ribbon:
      defaultConfig.ribbon?.filter((tab) => DESIGNER_TAB_IDS.has(tab.id)).map((tab) => ({
        ...tab,
        text: tab.id === "data" ? "Settings" : tab.text,
      })) ?? [],
  };
}

export function SpreadsheetEditor({
  snapshot,
  onSnapshotChange,
  onApiReady,
}: SpreadsheetEditorProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const spreadRef = useRef<SpreadsheetApi | null>(null);
  const designerRef = useRef<DesignerInstance | null>(null);
  const saveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const resizeFrameRef = useRef<number | null>(null);
  const initializedRef = useRef(false);
  const hydratingRef = useRef(false);
  const lastAppliedSnapshotRef = useRef<string | null>(null);

  const flushSnapshot = useCallback(() => {
    if (!onSnapshotChange || !spreadRef.current || hydratingRef.current) return;
    const nextSnapshot = JSON.stringify(spreadRef.current.toJSON());
    lastAppliedSnapshotRef.current = nextSnapshot;
    onSnapshotChange(nextSnapshot);
  }, [onSnapshotChange]);

  const applySnapshot = useCallback(async (nextSnapshot?: string) => {
    const spread = spreadRef.current;
    const designer = designerRef.current;
    if (!spread || !designer) return;

    const serialized = nextSnapshot ?? null;
    if (serialized === lastAppliedSnapshotRef.current) return;

    hydratingRef.current = true;
    const parsedSnapshot = parseSnapshot(nextSnapshot);
    const shouldPersistBlankWorkbook =
      parsedSnapshot === null || isLegacyUniverSnapshot(parsedSnapshot);

    try {
      if (parsedSnapshot && isSpreadSnapshot(parsedSnapshot)) {
        await spread.fromJSON(parsedSnapshot);
      }
    } catch (error) {
      console.warn("Failed to restore SpreadJS snapshot, using a blank workbook", error);
    } finally {
      configureWorkbook(spread);
      designer.activeRibbonTab("home");
      designer.refresh();
      hydratingRef.current = false;
      lastAppliedSnapshotRef.current = serialized;
    }

    if (shouldPersistBlankWorkbook) {
      flushSnapshot();
    }
  }, [flushSnapshot]);

  const scheduleSave = useCallback(() => {
    if (!onSnapshotChange || !spreadRef.current || hydratingRef.current) return;
    if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      flushSnapshot();
    }, SAVE_DEBOUNCE_MS);
  }, [flushSnapshot, onSnapshotChange]);

  useEffect(() => {
    if (!containerRef.current || initializedRef.current) return;
    initializedRef.current = true;

    const licenseKey = import.meta.env.VITE_SPREADJS_LICENSE_KEY;
    const designerLicenseKey =
      import.meta.env.VITE_SPREADJS_DESIGNER_LICENSE_KEY ?? licenseKey;
    const Designer = getDesignerNamespace();

    if (licenseKey) {
      GC.Spread.Sheets.LicenseKey = licenseKey;
    }
    if (designerLicenseKey) {
      Designer.LicenseKey = designerLicenseKey;
    }

    const designer = new Designer.Designer(
      containerRef.current,
      createDesignerConfig(),
      undefined,
      { sheetCount: 1 },
    );
    designerRef.current = designer;

    const spread = designer.getWorkbook() as SpreadsheetApi;
    spreadRef.current = spread;

    const handleWorkbookChange = () => {
      scheduleSave();
    };

    const handleContainerResize = () => {
      if (resizeFrameRef.current !== null) {
        cancelAnimationFrame(resizeFrameRef.current);
      }

      resizeFrameRef.current = requestAnimationFrame(() => {
        spread.invalidateLayout();
        spread.refresh();
        spread.getActiveSheet()?.invalidateLayout();
        spread.getActiveSheet()?.repaint();
        designer.refresh();
      });
    };

    const resizeObserver = new ResizeObserver(handleContainerResize);
    resizeObserver.observe(containerRef.current);

    void applySnapshot(snapshot).then(() => {
      onApiReady?.(spread);
      handleContainerResize();
    });

    spread.bind(GC.Spread.Sheets.Events.SheetChanged, handleWorkbookChange);
    spread.bind(GC.Spread.Sheets.Events.ClipboardChanged, handleWorkbookChange);
    spread.bind(GC.Spread.Sheets.Events.RangeChanged, handleWorkbookChange);
    spread.bind(GC.Spread.Sheets.Events.CellChanged, handleWorkbookChange);

    return () => {
      spread.unbind(GC.Spread.Sheets.Events.SheetChanged, handleWorkbookChange);
      spread.unbind(GC.Spread.Sheets.Events.ClipboardChanged, handleWorkbookChange);
      spread.unbind(GC.Spread.Sheets.Events.RangeChanged, handleWorkbookChange);
      spread.unbind(GC.Spread.Sheets.Events.CellChanged, handleWorkbookChange);
      resizeObserver.disconnect();
      if (resizeFrameRef.current !== null) {
        cancelAnimationFrame(resizeFrameRef.current);
      }
      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
      designer.destroy();
      designerRef.current = null;
      spreadRef.current = null;
      initializedRef.current = false;
    };
  }, [applySnapshot, onApiReady, scheduleSave, snapshot]);

  useEffect(() => {
    if (!initializedRef.current || !spreadRef.current) return;
    void applySnapshot(snapshot);
  }, [applySnapshot, snapshot]);

  return <div ref={containerRef} className="spread-designer-host" />;
}
