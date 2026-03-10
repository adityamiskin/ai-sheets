import { useQuery } from "convex/react";
import { api } from "../../../convex/_generated/api";
import type { Id } from "../../../convex/_generated/dataModel";
import {
  Queue,
  QueueItem,
  QueueItemContent,
  QueueItemIndicator,
  QueueList,
  QueueSection,
  QueueSectionContent,
  QueueSectionLabel,
  QueueSectionTrigger,
} from "@/components/ai-elements/queue";

interface TaskListProps {
  workbookId: Id<"workbooks"> | null;
}

export function TaskList({ workbookId }: TaskListProps) {
  const tasks = useQuery(
    api.tasks.list,
    workbookId ? { workbookId } : "skip"
  );

  if (!tasks || tasks.length === 0) return null;

  return (
    <div className="border-b border-[var(--linen)] px-4 py-2.5">
      <Queue className="gap-0 rounded-none border-0 bg-transparent px-0 py-0 shadow-none">
        <QueueSection defaultOpen>
          <QueueSectionTrigger className="rounded-none bg-transparent px-0 py-0 font-semibold text-[11px] uppercase tracking-[0.18em] text-[var(--umber)] hover:bg-transparent">
            <QueueSectionLabel count={tasks.length} label="tasks" />
          </QueueSectionTrigger>
          <QueueSectionContent>
            <QueueList className="mt-2 h-[176px]">
              {tasks.map((task, i) => (
                <QueueItem
                  key={task._id}
                  className="gap-0 rounded-none px-0 py-0.5 hover:bg-transparent"
                >
                  <div className="flex items-start gap-2">
                    <QueueItemIndicator completed={task.status === "completed"} />
                    <QueueItemContent
                      completed={task.status === "completed"}
                      className="line-clamp-none text-[13px] leading-5 text-[var(--graphite)]"
                    >
                      {i + 1}. {task.content}
                    </QueueItemContent>
                  </div>
                </QueueItem>
              ))}
            </QueueList>
          </QueueSectionContent>
        </QueueSection>
      </Queue>
    </div>
  );
}
