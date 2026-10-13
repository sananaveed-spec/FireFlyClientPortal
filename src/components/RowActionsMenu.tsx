"use client";

type RowActionsMenuProps = {
  onEdit: () => void;
};

export function RowActionsMenu({ onEdit }: RowActionsMenuProps) {
  return (
    <div className="row-menu">
      <button
        type="button"
        className="row-menu-trigger"
        aria-label="Edit row"
        title="Edit"
        onClick={onEdit}
      >
        ⋯
      </button>
    </div>
  );
}
