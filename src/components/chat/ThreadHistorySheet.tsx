import { useMemo, useState } from "react";
import { useQuery } from "convex/react";
import { api } from "../../../convex/_generated/api";
import type { Id } from "../../../convex/_generated/dataModel";
import {
  Sheet,
  SheetContent,
  SheetDescription,
  SheetHeader,
  SheetTitle,
} from "@/components/ui/sheet";
import { History, MessageSquarePlus, PanelLeftClose, Trash2 } from "lucide-react";

interface ThreadHistorySheetProps {
  workbookId: Id<"workbooks"> | null;
  activeThreadId: string | null;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onSelectThread: (threadId: string) => void;
  onCreateThread: () => Promise<string | null>;
  onDeleteThread: (threadId: string) => Promise<string | null>;
}

type ThreadItem = {
  threadId: string;
  title: string;
  createdAt: number;
  updatedAt: number;
};

type ThreadGroup = {
  label: string;
  threads: ThreadItem[];
};

function formatTimestamp(timestamp: number) {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(timestamp);
}

function getGroupLabel(timestamp: number, now: Date) {
  const target = new Date(timestamp);
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfTarget = new Date(target.getFullYear(), target.getMonth(), target.getDate());
  const dayDiff = Math.round((startOfToday.getTime() - startOfTarget.getTime()) / 86400000);

  if (dayDiff <= 0) return "Today";
  if (dayDiff === 1) return "Yesterday";
  if (dayDiff < 7) return "This week";
  if (dayDiff < 30) return "This month";
  return "Older";
}

function groupThreads(threads: ThreadItem[]) {
  const now = new Date();
  const groups = new Map<string, ThreadItem[]>();

  for (const thread of threads) {
    const label = getGroupLabel(thread.createdAt, now);
    const existing = groups.get(label) ?? [];
    existing.push(thread);
    groups.set(label, existing);
  }

  return ["Today", "Yesterday", "This week", "This month", "Older"]
    .map((label) => ({
      label,
      threads: groups.get(label) ?? [],
    }))
    .filter((group) => group.threads.length > 0);
}

export function ThreadHistorySheet({
  workbookId,
  activeThreadId,
  open,
  onOpenChange,
  onSelectThread,
  onCreateThread,
  onDeleteThread,
}: ThreadHistorySheetProps) {
  const threads = useQuery(
    api.chat.listThreadsForWorkbook,
    workbookId ? { workbookId } : "skip"
  );
  const [deletingThreadId, setDeletingThreadId] = useState<string | null>(null);

  const groupedThreads = useMemo<ThreadGroup[]>(
    () => groupThreads(threads ?? []),
    [threads]
  );

  return (
    <Sheet open={open} onOpenChange={onOpenChange}>
      <SheetContent
        side="left"
        showCloseButton={false}
        className="ramp-history-sheet"
        aria-label="Thread history"
      >
        <SheetHeader className="ramp-history-header">
          <div>
            <SheetTitle className="ramp-history-title">Thread history</SheetTitle>
            <SheetDescription className="ramp-history-description">
              Reopen prior chats or start a new thread.
            </SheetDescription>
          </div>
          <button
            type="button"
            className="ramp-history-icon-button"
            aria-label="Close thread history"
            onClick={() => onOpenChange(false)}
          >
            <PanelLeftClose className="h-4 w-4" />
          </button>
        </SheetHeader>

        <button
          type="button"
          className="ramp-history-new-thread"
          onClick={async () => {
            const threadId = await onCreateThread();
            if (threadId) onOpenChange(false);
          }}
          disabled={!workbookId}
        >
          <MessageSquarePlus className="h-4 w-4" />
          New thread
        </button>

        <div className="ramp-history-body">
          {!threads ? (
            <div className="ramp-history-empty">Loading thread history...</div>
          ) : groupedThreads.length === 0 ? (
            <div className="ramp-history-empty">
              <History className="h-4 w-4" />
              No saved threads yet.
            </div>
          ) : (
            groupedThreads.map((group) => (
              <section key={group.label} className="ramp-history-group">
                <div className="ramp-history-group-label">{group.label}</div>
                <div className="ramp-history-group-list">
                  {group.threads.map((thread) => {
                    const isActive = thread.threadId === activeThreadId;
                    return (
                      <div
                        key={thread.threadId}
                        className={`ramp-history-thread${isActive ? " is-active" : ""}`}
                      >
                        <button
                          type="button"
                          className="ramp-history-thread-main"
                          onClick={() => {
                            onSelectThread(thread.threadId);
                            onOpenChange(false);
                          }}
                        >
                          <span className="ramp-history-thread-title">{thread.title}</span>
                          <span className="ramp-history-thread-time">
                            {formatTimestamp(thread.updatedAt)}
                          </span>
                        </button>
                        <button
                          type="button"
                          className="ramp-history-thread-delete"
                          aria-label={`Delete thread ${thread.title}`}
                          disabled={deletingThreadId === thread.threadId}
                          onClick={async (event) => {
                            event.stopPropagation();
                            setDeletingThreadId(thread.threadId);
                            try {
                              await onDeleteThread(thread.threadId);
                            } finally {
                              setDeletingThreadId((current) =>
                                current === thread.threadId ? null : current
                              );
                            }
                          }}
                        >
                          <Trash2 className="h-4 w-4" />
                        </button>
                      </div>
                    );
                  })}
                </div>
              </section>
            ))
          )}
        </div>
      </SheetContent>
    </Sheet>
  );
}
