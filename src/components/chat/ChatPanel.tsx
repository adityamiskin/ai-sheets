import {
  Conversation,
  ConversationContent,
  ConversationEmptyState,
  ConversationScrollButton,
} from "@/components/ai-elements/conversation";
import { ChatMessage } from "./ChatMessage";
import { ChatInput } from "./ChatInput";
import { TaskList } from "./TaskList";
import { TooltipProvider } from "@/components/ui/tooltip";
import { useUIMessages } from "@convex-dev/agent/react";
import { api } from "../../../convex/_generated/api";
import type { Id } from "../../../convex/_generated/dataModel";
import { isToolUIPart, getToolName } from "ai";

interface ChatPanelProps {
  workbookId: Id<"workbooks"> | null;
  threadId: string | null;
  onSendMessage: (message: string) => void;
  isSending: boolean;
}

const SUGGESTED_PROMPTS = [
  "Create a monthly budget template",
  "Build a DCF valuation model",
  "Make a sales tracking dashboard",
  "Create a loan amortization schedule",
];

export function ChatPanel({
  workbookId,
  threadId,
  onSendMessage,
  isSending,
}: ChatPanelProps) {
  const { results, status } = useUIMessages(
    api.chat.listMessages,
    threadId ? { threadId } : "skip",
    { initialNumItems: 50, stream: true }
  );

  const isLoadingFirst = status === "LoadingFirstPage";

  const hasMessages = results && results.length > 0;

  return (
    <TooltipProvider>
      <div className="flex h-full min-h-0 w-full flex-col bg-[#fbfaf8]">
        <TaskList workbookId={workbookId} />

        <Conversation className="min-h-0 flex-1 overflow-hidden">
          {!hasMessages ? (
            <ConversationEmptyState
              className="items-start px-4 py-4 text-left"
              description="Open a file or start from scratch"
              title="Welcome to Zapsight Sheets"
            >
              <div className="space-y-4">
                <div>
                  <div className="mb-1 text-[18px] font-medium text-[var(--ink)]">Welcome to Zapsight Sheets</div>
                  <div className="text-sm text-[var(--soft-text)]">
                    Open a file or start from scratch
                  </div>
                </div>
                <div className="grid grid-cols-1 gap-2">
                  {SUGGESTED_PROMPTS.map((prompt) => (
                    <button
                      key={prompt}
                      onClick={() => onSendMessage(prompt)}
                      disabled={isSending}
                      className="rounded-[12px] border border-[var(--chat-border)] bg-white px-3 py-2.5 text-left text-[13px] text-[var(--ink)] transition hover:bg-[var(--chat-soft)] disabled:opacity-50"
                    >
                      {prompt}
                    </button>
                  ))}
                </div>
              </div>
            </ConversationEmptyState>
          ) : (
            <ConversationContent className="gap-4 px-4 py-4">
              {results.map((msg) => {
                const textContent = msg.text;
                const reasoningParts = msg.parts
                  .filter((part) => part.type === "reasoning")
                  .map((part) => ("text" in part ? part.text : ""))
                  .filter(Boolean);

                const toolParts = msg.parts.filter(isToolUIPart);
                const toolInvocations = toolParts.map((p) => ({
                  toolName: getToolName(p),
                  args: ("input" in p && typeof p.input === "object" && p.input !== null
                    ? p.input
                    : {}) as Record<string, unknown>,
                }));

                const isStreaming =
                  msg.status === "streaming" &&
                  msg.role !== "user";

                return (
                  <ChatMessage
                    key={msg.id}
                    role={msg.role === "user" ? "user" : "assistant"}
                    content={textContent}
                    reasoning={reasoningParts}
                    toolCalls={toolInvocations.length > 0 ? toolInvocations : undefined}
                    isStreaming={isStreaming}
                  />
                );
              })}
              {isLoadingFirst && (
                <div className="animate-pulse text-[13px] text-[var(--umber)]">
                  Loading messages...
                </div>
              )}
            </ConversationContent>
          )}
          <ConversationScrollButton className="bottom-3" />
        </Conversation>

        <ChatInput
          onSend={onSendMessage}
          disabled={isSending || !workbookId}
          isSending={isSending}
        />
      </div>
    </TooltipProvider>
  );
}
