import {
  Message,
  MessageAction,
  MessageActions,
  MessageContent,
} from "@/components/ai-elements/message";
import {
  Reasoning,
  ReasoningContent,
  ReasoningTrigger,
} from "@/components/ai-elements/reasoning";
import { Badge } from "@/components/ui/badge";
import { CheckIcon, CopyIcon } from "lucide-react";
import { useEffect, useState } from "react";
import { Streamdown } from "streamdown";

import "streamdown/styles.css";

interface ChatMessageProps {
  role: "user" | "assistant";
  content: string;
  toolCalls?: Array<{
    toolName: string;
    args: Record<string, unknown>;
  }>;
  reasoning?: string[];
  isStreaming?: boolean;
}

function formatToolName(toolName: string) {
  const normalized = toolName
    .split(".")
    .pop()
    ?.replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim() ?? toolName;

  return normalized.replace(/\b\w/g, (char) => char.toUpperCase());
}

export function ChatMessage({
  role,
  content,
  toolCalls,
  reasoning,
  isStreaming,
}: ChatMessageProps) {
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    if (!copied) return;
    const timer = setTimeout(() => setCopied(false), 1200);
    return () => clearTimeout(timer);
  }, [copied]);

  const handleCopy = async () => {
    if (typeof window === "undefined" || !navigator?.clipboard?.writeText) return;
    await navigator.clipboard.writeText(content ?? "");
    setCopied(true);
  };

  return (
    <Message
      from={role}
      className="max-w-full gap-2"
    >
      <MessageContent
        className={
          role === "assistant"
            ? "w-full gap-2 text-sm leading-6"
            : "gap-2 text-sm leading-6"
        }
      >
        {role === "assistant" && reasoning && reasoning.length > 0 && (
          <Reasoning
            className="mb-1 w-full"
            isStreaming={isStreaming}
          >
            <ReasoningTrigger className="gap-1 text-[12px] text-[#7d858f]" />
            <ReasoningContent className="mt-2 text-[12px] leading-5 text-[#7d858f]">
              {reasoning.join("\n\n")}
            </ReasoningContent>
          </Reasoning>
        )}

        {toolCalls && toolCalls.length > 0 && (
          <div className="flex flex-wrap gap-1.5">
            {toolCalls.map((tc, i) => (
              <Badge
                key={`${tc.toolName}-${i}`}
                variant="secondary"
                className="rounded-full border border-[var(--chat-border)] bg-[var(--chat-soft)] px-2 py-0.5 text-[11px] font-medium text-[#6f7882]"
              >
                {formatToolName(tc.toolName)}
              </Badge>
            ))}
          </div>
        )}

        {content ? (
          role === "assistant" ? (
            <Streamdown
              className="text-sm leading-6 text-[var(--graphite)] [&_a]:text-[var(--ink)] [&_a]:underline [&_a]:underline-offset-4 [&_code]:rounded-md [&_code]:bg-[var(--bone)] [&_code]:px-1 [&_code]:py-0.5 [&_pre]:overflow-x-auto [&_pre]:rounded-xl [&_pre]:border [&_pre]:border-[var(--linen)] [&_pre]:bg-[var(--bone)] [&_pre]:p-3 [&_ul]:list-disc [&_ul]:pl-5 [&_ol]:list-decimal [&_ol]:pl-5 [&_p]:mb-3 [&_p:last-child]:mb-0"
              isAnimating={isStreaming}
            >
              {content}
            </Streamdown>
          ) : (
            <div className="whitespace-pre-wrap break-words text-sm leading-6 text-[var(--graphite)]">
              {content}
            </div>
          )
        ) : isStreaming ? (
          <div className="animate-pulse text-sm text-[var(--umber)]">
            Thinking...
          </div>
        ) : null}
      </MessageContent>
      {Boolean(content?.trim()) && (
        <MessageActions className="ml-0">
          <MessageAction
            tooltip={copied ? "Copied" : "Copy"}
            label="Copy message"
            onClick={() => {
              void handleCopy();
            }}
          >
            {copied ? <CheckIcon className="size-4" /> : <CopyIcon className="size-4" />}
          </MessageAction>
        </MessageActions>
      )}
    </Message>
  );
}
