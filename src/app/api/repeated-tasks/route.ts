import { NextResponse } from "next/server";
import {
  normalizeClientRequestListsStore,
  promoteAllListsForNewMonth,
} from "@/lib/client-request-lists";
import {
  readClientRequestListsFile,
  readHolidays,
  readRepeatedTasks,
  writeClientRequestListsFile,
  writeHolidays,
  writeRepeatedTasks,
} from "@/lib/data-store";
import {
  ensureRepeatedTasks,
  getClientManagementNow,
  normalizeRepeatedTasks,
  ROLL_CLIENT_MANAGEMENT_MONTH_ID,
  runRepeatedTask,
  setChecklistItemReview,
  setClientManagementDemoMonth,
  type RepeatedTasksStore,
} from "@/lib/repeated-tasks";
import { requireEditor } from "@/lib/requireEditor";

/** Open BD5/BD9 action rows for the Client Management clock (demo or real). */
async function promoteRequestListsForClientManagementClock(
  tasks: RepeatedTasksStore,
  now: Date = new Date(),
) {
  const cmTask = tasks.tasks.find(
    (task) => task.id === ROLL_CLIENT_MANAGEMENT_MONTH_ID,
  );
  const clock = getClientManagementNow(cmTask?.demoMonth, now);
  const current = normalizeClientRequestListsStore(
    await readClientRequestListsFile(),
  );
  const promoted = promoteAllListsForNewMonth(current, clock);
  if (promoted.changed) {
    await writeClientRequestListsFile(promoted.store);
  }
}

export async function GET() {
  try {
    const data = await readRepeatedTasks();
    return NextResponse.json(data);
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to load repeated tasks.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}

type PostBody = {
  action?: "ensure" | "run" | "reviewChecklist" | "setDemoMonth";
  taskId?: string;
  itemId?: string;
  checked?: boolean;
  checkedBy?: string;
  demoMonth?: string | null;
};

/**
 * POST { action: "ensure" } — idempotent auto-run (Lookup open).
 * POST { action: "run", taskId } — admin recovery Run now.
 * POST { action: "setDemoMonth", demoMonth } — demo month for Client Management.
 * POST { action: "reviewChecklist", taskId, itemId, checked, checkedBy }
 */
export async function POST(request: Request) {
  try {
    const body = (await request.json().catch(() => ({}))) as PostBody;
    const action =
      body.action === "run"
        ? "run"
        : body.action === "reviewChecklist"
          ? "reviewChecklist"
          : body.action === "setDemoMonth"
            ? "setDemoMonth"
            : "ensure";

    if (action === "setDemoMonth") {
      const denied = await requireEditor(request);
      if (denied) {
        return denied;
      }
      const demoMonth =
        body.demoMonth === null || body.demoMonth === ""
          ? null
          : String(body.demoMonth ?? "").trim();
      if (demoMonth && !/^\d{4}-\d{2}$/.test(demoMonth)) {
        return NextResponse.json(
          { error: "demoMonth must be YYYY-MM or null." },
          { status: 400 },
        );
      }
      const [tasks, holidays] = await Promise.all([
        readRepeatedTasks(),
        readHolidays(),
      ]);
      const result = setClientManagementDemoMonth(
        tasks,
        holidays,
        demoMonth,
      );
      await writeRepeatedTasks(result.tasks);
      await promoteRequestListsForClientManagementClock(result.tasks);
      return NextResponse.json({
        ...result.tasks,
        ran: result.ran,
        message: result.message,
      });
    }

    if (action === "run") {
      const denied = await requireEditor(request);
      if (denied) {
        return denied;
      }
      const taskId = String(body.taskId ?? "").trim();
      if (!taskId) {
        return NextResponse.json(
          { error: "taskId is required for run." },
          { status: 400 },
        );
      }
      const [tasks, holidays] = await Promise.all([
        readRepeatedTasks(),
        readHolidays(),
      ]);
      const result = runRepeatedTask(taskId, tasks, holidays);
      await writeHolidays(result.holidays);
      await writeRepeatedTasks(result.tasks);
      if (taskId === ROLL_CLIENT_MANAGEMENT_MONTH_ID || result.ran) {
        await promoteRequestListsForClientManagementClock(result.tasks);
      }
      return NextResponse.json({
        ...result.tasks,
        ran: result.ran,
        message: result.message,
      });
    }

    if (action === "reviewChecklist") {
      const denied = await requireEditor(request);
      if (denied) {
        return denied;
      }
      const taskId = String(body.taskId ?? "").trim();
      const itemId = String(body.itemId ?? "").trim();
      const checkedBy = String(body.checkedBy ?? "").trim();
      if (!taskId || !itemId) {
        return NextResponse.json(
          { error: "taskId and itemId are required." },
          { status: 400 },
        );
      }
      const tasks = await readRepeatedTasks();
      const next = setChecklistItemReview(
        tasks,
        taskId,
        itemId,
        Boolean(body.checked),
        checkedBy,
      );
      await writeRepeatedTasks(next);
      return NextResponse.json({
        ...next,
        message: body.checked
          ? "Marked as checked."
          : "Marked as remaining.",
      });
    }

    const [tasks, holidays] = await Promise.all([
      readRepeatedTasks(),
      readHolidays(),
    ]);
    const result = ensureRepeatedTasks(tasks, holidays);
    if (result.ran) {
      await writeHolidays(result.holidays);
      await writeRepeatedTasks(result.tasks);
      await promoteRequestListsForClientManagementClock(result.tasks);
    } else {
      // Still persist nextPerformAt / normalize seed if file was missing fields.
      const normalized = normalizeRepeatedTasks(result.tasks);
      await writeRepeatedTasks(normalized);
    }
    return NextResponse.json({
      ...result.tasks,
      ran: result.ran,
      message: result.message,
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "Failed to process repeated tasks.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
