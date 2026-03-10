import { useAction, useMutation, useQuery } from "convex/react";
import { api } from "../../convex/_generated/api";
import { useCallback, useEffect, useState } from "react";
import type { Id } from "../../convex/_generated/dataModel";

/**
 * Manages chat thread creation and message sending.
 */
export function useChat(workbookId: Id<"workbooks"> | null) {
  const [threadId, setThreadId] = useState<string | null>(null);
  const [isSending, setIsSending] = useState(false);
  const existingThreadId = useQuery(
    api.chat.getThreadForWorkbook,
    workbookId ? { workbookId } : "skip"
  );

  const createChatThread = useMutation(api.chat.createChatThread);
  const sendMessageMutation = useMutation(api.chat.sendMessage);
  const deleteThreadAction = useAction(api.threadManagement.deleteThread);

  useEffect(() => {
    if (!existingThreadId) return;
    setThreadId((current) => current ?? existingThreadId);
  }, [existingThreadId]);

  const ensureThread = useCallback(async () => {
    if (threadId) return threadId;
    if (existingThreadId) {
      setThreadId(existingThreadId);
      return existingThreadId;
    }
    if (!workbookId) throw new Error("No workbook ID");
    const id = await createChatThread({ workbookId });
    setThreadId(id);
    return id;
  }, [threadId, existingThreadId, workbookId, createChatThread]);

  const createNewThread = useCallback(async () => {
    if (!workbookId) return null;
    const id = await createChatThread({ workbookId });
    setThreadId(id);
    return id;
  }, [workbookId, createChatThread]);

  const deleteThread = useCallback(
    async (targetThreadId: string) => {
      const result = await deleteThreadAction({ threadId: targetThreadId });

      if (threadId === targetThreadId) {
        if (result.nextThreadId) {
          setThreadId(result.nextThreadId);
          return result.nextThreadId;
        }

        if (workbookId) {
          const newThreadId = await createChatThread({ workbookId });
          setThreadId(newThreadId);
          return newThreadId;
        }

        setThreadId(null);
        return null;
      }

      return result.nextThreadId;
    },
    [createChatThread, deleteThreadAction, threadId, workbookId]
  );

  const sendMessage = useCallback(
    async (prompt: string) => {
      if (!workbookId) return;
      setIsSending(true);
      try {
        const tid = await ensureThread();
        await sendMessageMutation({
          threadId: tid,
          workbookId,
          prompt,
        });
      } finally {
        setIsSending(false);
      }
    },
    [workbookId, ensureThread, sendMessageMutation]
  );

  return {
    threadId: threadId ?? existingThreadId ?? null,
    isSending,
    sendMessage,
    ensureThread,
    setActiveThreadId: setThreadId,
    createNewThread,
    deleteThread,
  };
}
