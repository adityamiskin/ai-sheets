import { useMutation } from "convex/react";
import { api } from "../../convex/_generated/api";
import { useCallback, useState } from "react";
import type { Id } from "../../convex/_generated/dataModel";

/**
 * Manages chat thread creation and message sending.
 */
export function useChat(workbookId: Id<"workbooks"> | null) {
  const [threadId, setThreadId] = useState<string | null>(null);
  const [isSending, setIsSending] = useState(false);

  const getOrCreateThread = useMutation(api.chat.getOrCreateThread);
  const sendMessageMutation = useMutation(api.chat.sendMessage);

  const ensureThread = useCallback(async () => {
    if (threadId) return threadId;
    if (!workbookId) throw new Error("No workbook ID");
    const id = await getOrCreateThread({ workbookId });
    setThreadId(id);
    return id;
  }, [threadId, workbookId, getOrCreateThread]);

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
    threadId,
    isSending,
    sendMessage,
    ensureThread,
  };
}
