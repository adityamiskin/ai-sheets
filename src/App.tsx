import { useCallback, useEffect, useRef, useState } from "react";
import * as GC from "@mescius/spread-sheets";
import { useMutation, useQuery } from "convex/react";
import { useNavigate, useParams } from "@tanstack/react-router";
import { api } from "../convex/_generated/api";
import {
  SpreadsheetEditor,
  type SpreadsheetApi,
} from "./components/spreadsheet/SpreadsheetEditor";
import { ChatPanel } from "./components/chat/ChatPanel";
import { ThreadHistorySheet } from "./components/chat/ThreadHistorySheet";
import { DataSourcesSheet } from "./components/chat/DataSourcesSheet";
import {
  ResizableHandle,
  ResizablePanel,
  ResizablePanelGroup,
} from "./components/ui/resizable";
import { Button } from "./components/ui/button";
import { useWorkbook } from "./hooks/useWorkbook";
import { useChat } from "./hooks/useChat";
import { useOperationProcessor } from "./hooks/useOperationProcessor";
import { Database, Download, Menu } from "lucide-react";

export default function App() {
  const navigate = useNavigate();
  const params = useParams({ strict: false });
  const routeThreadId =
    typeof params.threadId === "string" ? params.threadId : null;
  const { workbookId, workbook, initialize, saveSnapshot } = useWorkbook();
  const {
    threadId,
    isSending,
    sendMessage,
    setActiveThreadId,
    createNewThread,
    deleteThread,
  } = useChat(workbookId, routeThreadId);
  const activeThread = useQuery(
    api.chat.getThreadState,
    threadId ? { threadId } : "skip"
  );
  const updateThreadSnapshot = useMutation(api.chat.updateThreadSnapshot);
  const [spread, setSpread] = useState<SpreadsheetApi | null>(null);
  const [isHistoryOpen, setIsHistoryOpen] = useState(false);
  const [isDataSourcesOpen, setIsDataSourcesOpen] = useState(false);
  const initializedRef = useRef(false);
  const saveChainRef = useRef(Promise.resolve());

  // Initialize workbook on mount
  useEffect(() => {
    if (initializedRef.current) return;
    initializedRef.current = true;
    initialize();
  }, [initialize]);

  useEffect(() => {
    if (!threadId) return;
    if (routeThreadId === threadId) return;

    void navigate({
      to: "/$threadId",
      params: { threadId },
      replace: routeThreadId === null,
    });
  }, [navigate, routeThreadId, threadId]);

  const isWorkbookLoading = workbookId === null || workbook === undefined;
  const isThreadSnapshotLoading = threadId !== null && activeThread === undefined;
  const currentSnapshot =
    threadId && activeThread ? activeThread.snapshot : workbook?.snapshot;
  const currentSnapshotVersion =
    threadId && activeThread ? activeThread.updatedAt : workbook?.updatedAt;

  const handleApiReady = useCallback((api: SpreadsheetApi) => {
    setSpread(api);
  }, []);

  const handleSnapshotChange = useCallback(
    (snapshot: string) => {
      saveChainRef.current = saveChainRef.current
        .catch(() => undefined)
        .then(async () => {
          await saveSnapshot(snapshot);
          if (threadId) {
            await updateThreadSnapshot({ threadId, snapshot });
          }
        });

      return saveChainRef.current;
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

  useOperationProcessor(workbookId, spread);

  const handleSelectThread = useCallback(
    (nextThreadId: string) => {
      setActiveThreadId(nextThreadId);
      void navigate({
        to: "/$threadId",
        params: { threadId: nextThreadId },
      });
    },
    [navigate, setActiveThreadId]
  );

  const handleCreateThread = useCallback(async () => {
    const nextThreadId = await createNewThread();
    if (nextThreadId) {
      void navigate({
        to: "/$threadId",
        params: { threadId: nextThreadId },
      });
    }
    return nextThreadId;
  }, [createNewThread, navigate]);

  const handleDeleteThread = useCallback(
    async (targetThreadId: string) => {
      const nextThreadId = await deleteThread(targetThreadId);
      if (nextThreadId) {
        void navigate({
          to: "/$threadId",
          params: { threadId: nextThreadId },
          replace: true,
        });
      } else {
        void navigate({
          to: "/",
          replace: true,
        });
      }
      return nextThreadId;
    },
    [deleteThread, navigate]
  );

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
            disabled={!workbookId}
            onClick={() => setIsDataSourcesOpen(true)}
          >
            <Database className="h-3.5 w-3.5" />
            Sources
          </Button>
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
          onSelectThread={handleSelectThread}
          onCreateThread={handleCreateThread}
          onDeleteThread={handleDeleteThread}
        />

        <DataSourcesSheet
          workbookId={workbookId}
          open={isDataSourcesOpen}
          onOpenChange={setIsDataSourcesOpen}
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
              ) : isThreadSnapshotLoading ? (
                <div className="ramp-loading-state">
                  Loading thread snapshot...
                </div>
              ) : (
                <SpreadsheetEditor
                  key={workbookId ?? "loading"}
                  snapshot={currentSnapshot}
                  snapshotVersion={currentSnapshotVersion}
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
            className="ramp-chat-pane text-sm"
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
