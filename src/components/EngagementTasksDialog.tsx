"use client";

import type { ReactNode } from "react";

export type EngagementTasksDialogProps = {
  title: string;
  titleId: string;
  intro: ReactNode;
  emptyHint: ReactNode;
  statusLabel: string;
  itemCount: number;
  saving: boolean;
  error: string | null;
  canEdit: boolean;
  onClose: () => void;
  children: ReactNode;
  addRow?: ReactNode;
};

/** Shared checklist dialog shell for One-Time / Annual / Monthly Tasks. */
export function EngagementTasksDialog({
  title,
  titleId,
  intro,
  emptyHint,
  statusLabel,
  itemCount,
  saving,
  error,
  canEdit,
  onClose,
  children,
  addRow,
}: EngagementTasksDialogProps) {
  return (
    <div className="dialog-backdrop" role="presentation" onClick={onClose}>
      <div
        className="dialog-panel dialog-panel--wide one-time-tasks-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        onClick={(event) => event.stopPropagation()}
      >
        <div className="dialog-header">
          <div className="one-time-tasks-title-row">
            <h2 id={titleId} className="dialog-title">
              {title}
            </h2>
            {itemCount > 0 && statusLabel ? (
              <span
                className={`one-time-tasks-status one-time-tasks-status--${statusLabel
                  .toLowerCase()
                  .replace(/\s+/g, "-")}`}
              >
                {statusLabel}
              </span>
            ) : null}
          </div>
          <button
            type="button"
            className="dialog-close"
            aria-label="Close"
            disabled={saving}
            onClick={onClose}
          >
            ×
          </button>
        </div>
        <p className="one-time-tasks-intro">{intro}</p>
        {saving && itemCount === 0 ? (
          <p className="one-time-tasks-empty">Loading tasks…</p>
        ) : itemCount === 0 ? (
          <div className="one-time-tasks-empty">{emptyHint}</div>
        ) : (
          <ul className="one-time-tasks-list">{children}</ul>
        )}
        {canEdit && addRow ? addRow : null}
        {error ? (
          <p className="form-message error" role="alert">
            {error}
          </p>
        ) : null}
        <div className="dialog-actions">
          <button
            type="button"
            className="button primary"
            disabled={saving}
            onClick={onClose}
          >
            Done
          </button>
        </div>
      </div>
    </div>
  );
}
