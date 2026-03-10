import {
  PromptInput,
  PromptInputBody,
  PromptInputFooter,
  PromptInputSubmit,
  PromptInputTextarea,
  PromptInputTools,
} from "@/components/ai-elements/prompt-input";
import type { ChatStatus } from "ai";

interface ChatInputProps {
  onSend: (message: string) => void;
  disabled?: boolean;
  placeholder?: string;
  isSending?: boolean;
}

export function ChatInput({
  onSend,
  disabled = false,
  placeholder = "Ask Zapsight Sheets anything...",
  isSending = false,
}: ChatInputProps) {
  return (
    <div className="border-t border-[var(--chat-border)] p-3">
      <PromptInput
        className="w-full"
        onSubmit={({ text }) => {
          const trimmed = text.trim();
          if (!trimmed || disabled) {
            throw new Error("skip-empty-submit");
          }
          onSend(trimmed);
        }}
      >
        <PromptInputBody>
          <PromptInputTextarea
            className="min-h-[42px] max-h-[200px] text-[13px] leading-6"
            disabled={disabled}
            placeholder={placeholder}
          />
        </PromptInputBody>
        <PromptInputFooter className="justify-end px-2.5 pb-2">
          <PromptInputTools />
          <PromptInputSubmit
            className="rounded-full border border-[var(--chat-border)] bg-[var(--chat-send)] text-white hover:bg-[var(--ink)] disabled:bg-[var(--chat-send)]/60"
            disabled={disabled}
            status={(isSending ? "submitted" : "ready") as ChatStatus}
          />
        </PromptInputFooter>
      </PromptInput>
    </div>
  );
}
