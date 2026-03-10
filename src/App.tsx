import { useCallback, useEffect, useRef, useState } from "react";
import * as GC from "@mescius/spread-sheets";
import { useMutation, useQuery } from "convex/react";
import { api } from "../convex/_generated/api";
import {
  SpreadsheetEditor,
  type SpreadsheetApi,
} from "./components/spreadsheet/SpreadsheetEditor";
import { ChatPanel } from "./components/chat/ChatPanel";
import { ThreadHistorySheet } from "./components/chat/ThreadHistorySheet";
import {
  ResizableHandle,
  ResizablePanel,
  ResizablePanelGroup,
} from "./components/ui/resizable";
import { Button } from "./components/ui/button";
import { useWorkbook } from "./hooks/useWorkbook";
import { useChat } from "./hooks/useChat";
import { useOperationProcessor } from "./hooks/useOperationProcessor";
import { Download, Menu } from "lucide-react";
export default function App() {
  const { workbookId, workbook, initialize, saveSnapshot, loadSnapshot } = useWorkbook();
  const {
    threadId,
    isSending,
    sendMessage,
    setActiveThreadId,
    createNewThread,
    deleteThread,
  } = useChat(workbookId);
  const activeThread = useQuery(
    api.chat.getThreadState,
    threadId ? { threadId } : "skip"
  );
  const updateThreadSnapshot = useMutation(api.chat.updateThreadSnapshot);
  const [spread, setSpread] = useState<SpreadsheetApi | null>(null);
  const [isHistoryOpen, setIsHistoryOpen] = useState(false);
  const initializedRef = useRef(false);
  const pendingThreadHydrationRef = useRef<string | null>(null);

  // Initialize workbook on mount
  useEffect(() => {
    if (initializedRef.current) return;
    initializedRef.current = true;
    initialize();
  }, [initialize]);

  useEffect(() => {
    pendingThreadHydrationRef.current = threadId;
  }, [threadId]);

  useEffect(() => {
    if (!threadId || !activeThread) return;
    if (pendingThreadHydrationRef.current !== threadId) return;

    pendingThreadHydrationRef.current = null;
    void loadSnapshot(activeThread.snapshot);
  }, [activeThread, loadSnapshot, threadId]);

  const isWorkbookLoading = workbookId === null || workbook === undefined;

  const handleApiReady = useCallback((api: SpreadsheetApi) => {
    setSpread(api);
  }, []);

  const handleSnapshotChange = useCallback(
    async (snapshot: string) => {
      await saveSnapshot(snapshot);
      if (threadId) {
        await updateThreadSnapshot({ threadId, snapshot });
      }
    },
    [saveSnapshot, threadId, updateThreadSnapshot]
  );

  const handleDownload = useCallback(() => {
    if (!spread) return;

    const workbookName = (workbook?.name ?? "Book1.xlsx").replace(/\.[^.]+$/, "");
    const fileName = `${workbookName}.xlsx`;

    (
      spread as unknown as {
        export: (
          success: (blob: Blob) => void,
          error?: (error: unknown) => void,
          options?: { fileType: unknown; includeBindingSource?: boolean },
        ) => void;
      }
    ).export(
      (blob) => {
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = fileName;
        link.click();
        URL.revokeObjectURL(url);
      },
      (error) => {
        console.error("Failed to export workbook", error);
      },
      {
        fileType: GC.Spread.Sheets.FileType.excel,
        includeBindingSource: true,
      }
    );
  }, [spread, workbook?.name]);

  useOperationProcessor(workbookId, spread, handleSnapshotChange);

  return (
    <div className="ramp-app-shell">
      <header className="ramp-topbar">
        <div className="ramp-header-meta">
          <button
            className="ramp-icon-button"
            type="button"
            aria-label="Open thread history"
            aria-expanded={isHistoryOpen}
            onClick={() => setIsHistoryOpen(true)}
          >
            <Menu className="h-4 w-4" />
          </button>
          <div className="ramp-wordmark">Zapsight Sheets</div>
          <div className="ramp-workbook-name">{workbook?.name ?? "Book1.xlsx"}</div>
        </div>

        <div className="ramp-topbar-actions">
          <Button
            type="button"
            variant="ghost"
            size="sm"
            disabled={!spread || isWorkbookLoading}
            onClick={handleDownload}
          >
            <Download className="h-3.5 w-3.5" />
            Download
          </Button>
          <div className="ramp-avatar">AI</div>
        </div>
      </header>

      <main className="ramp-workspace">
        <ThreadHistorySheet
          workbookId={workbookId}
          activeThreadId={threadId}
          open={isHistoryOpen}
          onOpenChange={setIsHistoryOpen}
          onSelectThread={setActiveThreadId}
          onCreateThread={createNewThread}
          onDeleteThread={deleteThread}
        />

        <ResizablePanelGroup
          orientation="horizontal"
          className="ramp-panel-group"
        >
          <ResizablePanel
            defaultSize="72%"
            minSize="45%"
            className="ramp-editor-pane"
          >
            <div className="ramp-sheet-stage">
              {isWorkbookLoading ? (
                <div className="ramp-loading-state">
                  Loading workbook...
                </div>
              ) : (
                <SpreadsheetEditor
                  key={workbookId ?? "loading"}
                  snapshot={workbook?.snapshot}
                  onSnapshotChange={handleSnapshotChange}
                  onApiReady={handleApiReady}
                />
              )}
            </div>
          </ResizablePanel>

          <ResizableHandle
            withHandle
            className="ramp-resizable-handle"
          />

          <ResizablePanel
            defaultSize="28%"
            minSize="24%"
            maxSize="45%"
            className="ramp-chat-pane"
          >
            <ChatPanel
              workbookId={workbookId}
              threadId={threadId}
              onSendMessage={sendMessage}
              isSending={isSending}
            />
          </ResizablePanel>
        </ResizablePanelGroup>
      </main>
    </div>
  );
}
