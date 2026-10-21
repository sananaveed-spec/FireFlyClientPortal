"use client";

import { useEffect, useMemo, useState } from "react";
import { useAccess } from "@/auth/accessContext";
import { portalFetch } from "@/lib/portal-api";
import {
  CLIENT_MANAGEMENT_DEMO_MONTH_EVENT,
  ROLL_CLIENT_MANAGEMENT_MONTH_ID,
  type ClientManagementBdRefresh,
  type RepeatedTask,
  type RepeatedTaskStatus,
  type RepeatedTasksStore,
} from "@/lib/repeated-tasks";

function formatStamp(value: string | null): string {
  if (!value) {
    return "—";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }
  return date.toLocaleString("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });
}

function statusLabel(status: RepeatedTaskStatus): string {
  switch (status) {
    case "idle":
      return "Idle";
    case "due":
      return "Due";
    case "running":
      return "Running";
    case "done":
      return "Done";
    case "failed":
      return "Failed";
    default:
      return status;
  }
}

type EnsureResponse = RepeatedTasksStore & {
  error?: string;
  ran?: boolean;
  message?: string;
};

type RepeatedTasksPanelProps = {
  onOpenHolidays?: () => void;
};

export function RepeatedTasksPanel({
  onOpenHolidays,
}: RepeatedTasksPanelProps) {
  const { canEdit } = useAccess();
  const [tasks, setTasks] = useState<RepeatedTask[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [selectedBdId, setSelectedBdId] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [runningId, setRunningId] = useState<string | null>(null);
  const [demoMonthDraft, setDemoMonthDraft] = useState("");
  const [savingDemoMonth, setSavingDemoMonth] = useState(false);

  const selected = useMemo(
    () => tasks.find((task) => task.id === selectedId) ?? tasks[0] ?? null,
    [tasks, selectedId],
  );

  const isMonthlyClientTask =
    selected?.id === ROLL_CLIENT_MANAGEMENT_MONTH_ID;

  const selectedBd: ClientManagementBdRefresh | null = useMemo(() => {
    if (!isMonthlyClientTask || !selected) {
      return null;
    }
    const slots = selected.bdSlots ?? [];
    return (
      slots.find((slot) => slot.id === selectedBdId) ?? slots[0] ?? null
    );
  }, [isMonthlyClientTask, selected, selectedBdId]);

  useEffect(() => {
    if (!isMonthlyClientTask || !selected?.bdSlots?.length) {
      return;
    }
    setSelectedBdId((current) => {
      if (current && selected.bdSlots.some((slot) => slot.id === current)) {
        return current;
      }
      return selected.bdSlots[0]?.id ?? null;
    });
  }, [isMonthlyClientTask, selected]);

  useEffect(() => {
    if (!isMonthlyClientTask || !selected) {
      return;
    }
    setDemoMonthDraft(selected.demoMonth ?? "");
  }, [isMonthlyClientTask, selected?.id, selected?.demoMonth]);

  useEffect(() => {
    let cancelled = false;

    async function loadAndEnsure() {
      try {
        const response = await portalFetch("/api/repeated-tasks", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ action: "ensure" }),
        });
        const data = (await response.json()) as EnsureResponse;
        if (!response.ok) {
          throw new Error(data.error || "Failed to load repeated tasks.");
        }
        if (!cancelled) {
          setTasks(data.tasks ?? []);
          setSelectedId((current) => current ?? data.tasks?.[0]?.id ?? null);
          setInfo(data.ran ? data.message || "Automatic task completed." : null);
          setError(null);
        }
      } catch (loadError) {
        if (!cancelled) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load repeated tasks.",
          );
        }
      } finally {
        if (!cancelled) {
          setReady(true);
        }
      }
    }

    void loadAndEnsure();
    return () => {
      cancelled = true;
    };
  }, []);

  async function handleSetDemoMonth(demoMonth: string | null) {
    setSavingDemoMonth(true);
    setError(null);
    setInfo(null);
    try {
      const response = await portalFetch("/api/repeated-tasks", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "setDemoMonth", demoMonth }),
      });
      const data = (await response.json()) as EnsureResponse;
      if (!response.ok) {
        throw new Error(data.error || "Failed to set demo month.");
      }
      setTasks(data.tasks ?? []);
      setInfo(data.message || "Demo month updated.");
      window.dispatchEvent(new Event(CLIENT_MANAGEMENT_DEMO_MONTH_EVENT));
    } catch (demoError) {
      setError(
        demoError instanceof Error
          ? demoError.message
          : "Failed to set demo month.",
      );
    } finally {
      setSavingDemoMonth(false);
    }
  }

  async function handleRunNow(taskId: string) {
    setRunningId(taskId);
    setError(null);
    setInfo(null);
    try {
      const response = await portalFetch("/api/repeated-tasks", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "run", taskId }),
      });
      const data = (await response.json()) as EnsureResponse;
      if (!response.ok) {
        throw new Error(data.error || "Failed to run task.");
      }
      setTasks(data.tasks ?? []);
      setInfo(data.message || "Task finished.");
      setSelectedId(taskId);
    } catch (runError) {
      setError(
        runError instanceof Error ? runError.message : "Failed to run task.",
      );
    } finally {
      setRunningId(null);
    }
  }

  if (!ready) {
    return <p className="table-empty">Loading…</p>;
  }

  return (
    <div className="lookup-layout repeated-tasks-panel">
      <aside className="lookup-fields" aria-label="Repeated tasks">
        {tasks.length === 0 ? (
          <p className="lookup-hint">No tasks.</p>
        ) : (
          tasks.map((task) => (
            <button
              key={task.id}
              type="button"
              className={
                selected?.id === task.id
                  ? "lookup-field-link active"
                  : "lookup-field-link"
              }
              onClick={() => setSelectedId(task.id)}
            >
              <span>{task.name}</span>
              <span
                className={`repeated-task-status repeated-task-status--${task.status}`}
              >
                {statusLabel(task.status)}
              </span>
            </button>
          ))
        )}
      </aside>

      <div className="lookup-items repeated-tasks-detail">
        <div className="repeated-tasks-status" aria-live="polite">
          {error ? (
            <p className="form-message error">{error}</p>
          ) : info ? (
            <p className="form-message success">{info}</p>
          ) : null}
        </div>

        {!selected ? (
          <p className="table-empty">Select a repeated task.</p>
        ) : (
          <>
            <div className="email-template-header">
              <div className="email-template-header__top">
                <h3 className="section-title">{selected.name}</h3>
                {canEdit ? (
                  <div className="lookup-actions">
                    <button
                      type="button"
                      className="button secondary-solid"
                      disabled={runningId === selected.id}
                      onClick={() => void handleRunNow(selected.id)}
                    >
                      {runningId === selected.id ? "Running…" : "Run now"}
                    </button>
                  </div>
                ) : null}
              </div>
              <p className="lookup-hint">{selected.description}</p>
            </div>

            <div className="email-template-fields">
              <div className="field">
                <span className="field-label">Last performed</span>
                <div className="field-input field-input--readonly">
                  {formatStamp(selected.lastPerformedAt)}
                </div>
              </div>
              <div className="field">
                <span className="field-label">Next perform</span>
                <div className="field-input field-input--readonly">
                  {formatStamp(selected.nextPerformAt)}
                </div>
              </div>
              <div className="field">
                <span className="field-label">Current status</span>
                <div className="field-input field-input--readonly">
                  <span
                    className={`repeated-task-status repeated-task-status--${selected.status}`}
                  >
                    {statusLabel(selected.status)}
                  </span>
                </div>
              </div>
              {isMonthlyClientTask ? (
                <div className="field">
                  <span className="field-label">Active month</span>
                  <div className="field-input field-input--readonly">
                    {selected.lastTargetMonth
                      ? `${selected.lastTargetMonth}${
                          selected.demoMonth ? " (demo)" : ""
                        }`
                      : "—"}
                  </div>
                </div>
              ) : null}
              {isMonthlyClientTask && canEdit ? (
                <div className="field field--wide">
                  <span className="field-label">Demo month</span>
                  <div className="repeated-tasks-demo-month">
                    <input
                      className="field-input"
                      type="month"
                      value={demoMonthDraft}
                      onChange={(event) =>
                        setDemoMonthDraft(event.target.value)
                      }
                    />
                    <button
                      type="button"
                      className={
                        savingDemoMonth
                          ? "button primary is-busy"
                          : "button primary"
                      }
                      disabled={!demoMonthDraft}
                      aria-busy={savingDemoMonth}
                      onClick={() => {
                        if (savingDemoMonth || !demoMonthDraft) {
                          return;
                        }
                        void handleSetDemoMonth(demoMonthDraft);
                      }}
                    >
                      {savingDemoMonth
                        ? "Applying…"
                        : "Apply to Client Management"}
                    </button>
                    <button
                      type="button"
                      className="button secondary"
                      disabled={savingDemoMonth || !selected.demoMonth}
                      onClick={() => void handleSetDemoMonth(null)}
                    >
                      Clear demo
                    </button>
                  </div>
                  <span className="lookup-hint">
                    Pick any month to preview BD dates in Client Management.
                    Clear to return to the real calendar month.
                  </span>
                </div>
              ) : null}
              {selected.lastError ? (
                <div className="field field--wide">
                  <span className="field-label">Last error</span>
                  <div className="field-input field-input--readonly">
                    {selected.lastError}
                  </div>
                </div>
              ) : null}
            </div>

            {isMonthlyClientTask ? (
              <div className="lookup-layout repeated-tasks-bd-layout">
                <aside className="lookup-fields" aria-label="Business days">
                  {(selected.bdSlots ?? []).length === 0 ? (
                    <p className="lookup-hint">
                      No BD slots yet. Run the monthly refresh.
                    </p>
                  ) : (
                    selected.bdSlots.map((slot) => (
                      <button
                        key={slot.id}
                        type="button"
                        className={
                          selectedBd?.id === slot.id
                            ? "lookup-field-link active"
                            : "lookup-field-link"
                        }
                        onClick={() => setSelectedBdId(slot.id)}
                      >
                        <span>{slot.label}</span>
                      </button>
                    ))
                  )}
                </aside>
                <div className="lookup-items repeated-tasks-bd-detail">
                  {!selectedBd ? (
                    <p className="table-empty">Select a business day.</p>
                  ) : (
                    <div className="email-template-fields">
                      <div className="field field--wide">
                        <span className="field-label">Business day</span>
                        <div className="field-input field-input--readonly">
                          {selectedBd.label}
                        </div>
                      </div>
                      <div className="field field--wide">
                        <span className="field-label">
                          Date this month
                        </span>
                        <div className="field-input field-input--readonly">
                          {selectedBd.businessDateLabel}
                        </div>
                      </div>
                      <div className="field field--wide">
                        <span className="field-label">
                          Default template list
                        </span>
                        <div className="field-input field-input--readonly">
                          {selectedBd.defaultTemplate}
                        </div>
                      </div>
                      <div className="field field--wide">
                        <span className="field-label">Due date rule</span>
                        <div className="field-input field-input--readonly">
                          {selectedBd.deadlineRule}
                        </div>
                      </div>
                      <div className="field field--wide">
                        <span className="field-label">Last refreshed</span>
                        <div className="field-input field-input--readonly">
                          {formatStamp(selectedBd.refreshedAt)}
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            ) : (
              <div className="field field--wide">
                <span className="field-label">Checklist</span>
                <div className="field-input field-input--readonly repeated-task-checklist">
                  {selected.checklist.length === 0 ? (
                    <p className="lookup-hint">
                      No checklist entries yet. The system fills these when the
                      task runs.
                    </p>
                  ) : (
                    <ul>
                      {selected.checklist.map((item) => (
                        <li key={item.id}>
                          <div className="repeated-task-checklist__row">
                            <span
                              className={
                                item.manual
                                  ? "repeated-task-check"
                                  : item.done
                                    ? "repeated-task-check done"
                                    : "repeated-task-check"
                              }
                              aria-hidden
                            >
                              {item.manual ? "✎" : item.done ? "✓" : "○"}
                            </span>
                            <div className="repeated-task-checklist__body">
                              <div className="repeated-task-checklist__title-row">
                                <strong>{item.label}</strong>
                                {item.manual ? (
                                  <span className="repeated-task-status repeated-task-status--manual">
                                    Manual
                                  </span>
                                ) : null}
                              </div>
                              {item.detail ? (
                                <p className="lookup-hint">{item.detail}</p>
                              ) : null}
                              <p className="lookup-hint">
                                {formatStamp(item.performedAt)}
                              </p>
                              {item.manual ? (
                                <div
                                  className="inline-notice inline-notice--warning repeated-task-lunar-notice"
                                  role="status"
                                >
                                  <p className="inline-notice__text">
                                    <strong>Action needed:</strong> Lunar
                                    holidays must be updated manually in
                                    Holidays. The auto year-shift is only a
                                    placeholder.
                                  </p>
                                  {onOpenHolidays ? (
                                    <button
                                      type="button"
                                      className="button secondary-solid"
                                      onClick={onOpenHolidays}
                                    >
                                      Open Holidays
                                    </button>
                                  ) : null}
                                </div>
                              ) : null}
                            </div>
                          </div>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
