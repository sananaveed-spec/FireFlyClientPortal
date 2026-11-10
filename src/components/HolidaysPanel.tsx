"use client";

import { useMsal } from "@azure/msal-react";
import { useEffect, useMemo, useState, type FormEvent } from "react";
import { createPortal } from "react-dom";
import { useAccess } from "@/auth/accessContext";
import { SearchableSelect } from "@/components/SearchableSelect";
import {
  getActorFromAccount,
  logPortalAction,
} from "@/lib/history-client";
import { isLunarHolidayName } from "@/lib/holiday-rules";
import {
  createHolidayEntry,
  formatHolidayDate,
  getHolidayRegionLabel,
  getHolidayRegionOptions,
  getHolidayWeekday,
  HOLIDAY_COUNTRIES,
  HOLIDAY_RELIGION_ALL,
  HOLIDAY_RELIGIONS,
  HOLIDAY_STATE_ALL,
  holidayUsesRegion,
  holidayUsesReligion,
  isHolidayManuallySet,
  normalizeHolidayDate,
  normalizeHolidayReligion,
  normalizeHolidayState,
  type HolidayCountryId,
  type HolidayEntry,
  type HolidayReligion,
  type HolidaysStore,
} from "@/lib/holidays";
import {
  buildPakistanHolidayEmailPreview,
  isPakistanHolidayEmailSender,
  PAKISTAN_HOLIDAY_EMAIL_SENDER,
  type PakistanHolidayEmailPreviewRow,
} from "@/lib/pakistan-holiday-emails";
import { portalFetch } from "@/lib/portal-api";

type EditorState = {
  mode: "add" | "edit";
  holidayId?: string;
  name: string;
  date: string;
  state: string;
  religion: HolidayReligion;
};

async function persistHolidays(data: HolidaysStore) {
  const response = await portalFetch("/api/holidays", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(data),
  });
  const payload = (await response.json()) as HolidaysStore & {
    error?: string;
  };
  if (!response.ok) {
    throw new Error(payload.error || "Failed to save holidays.");
  }
  return payload;
}

export function HolidaysPanel() {
  const { instance, accounts } = useMsal();
  const { canEdit, email: portalEmail } = useAccess();
  const actor = getActorFromAccount(accounts[0]);
  const canSendPakistanHolidayEmails =
    canEdit && isPakistanHolidayEmailSender(portalEmail || "");
  const [store, setStore] = useState<HolidaysStore>({ countries: [] });
  const [activeCountryId, setActiveCountryId] =
    useState<HolidayCountryId>("pakistan");
  const [editor, setEditor] = useState<EditorState | null>(null);
  const [ready, setReady] = useState(false);
  const [mounted, setMounted] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [info, setInfo] = useState<string | null>(null);
  const [lunarNoticeDismissed, setLunarNoticeDismissed] = useState(false);
  const [emailStatusByHolidayId, setEmailStatusByHolidayId] = useState<
    Record<string, "Scheduled" | "Sent" | "—">
  >({});
  const [clientEmailStatusByHolidayId, setClientEmailStatusByHolidayId] =
    useState<Record<string, "Scheduled" | "Sent" | "—">>({});
  const [upcomingHolidayEmails, setUpcomingHolidayEmails] = useState<
    PakistanHolidayEmailPreviewRow[]
  >([]);
  const [schedulingHolidayEmails, setSchedulingHolidayEmails] = useState(false);

  const showRegionColumn = holidayUsesRegion(activeCountryId);
  const showReligionColumn = holidayUsesReligion(activeCountryId);
  const showEmailStatusColumn = activeCountryId === "pakistan";
  const regionLabel = getHolidayRegionLabel(activeCountryId);
  const regionOptions = useMemo(
    () => getHolidayRegionOptions(activeCountryId),
    [activeCountryId],
  );
  const religionOptions = useMemo(() => [...HOLIDAY_RELIGIONS], []);
  const currentYear = new Date().getFullYear();
  const pakistanHolidays = useMemo(
    () =>
      store.countries.find((country) => country.id === "pakistan")?.holidays ??
      [],
    [store.countries],
  );
  /** Local preview so viewers (e.g. sana@epsfresno.com) always see upcoming sends. */
  const localUpcomingHolidayEmails = useMemo(
    () => buildPakistanHolidayEmailPreview(pakistanHolidays, { entries: [] }),
    [pakistanHolidays],
  );
  const displayUpcomingHolidayEmails =
    upcomingHolidayEmails.length > 0
      ? upcomingHolidayEmails
      : showEmailStatusColumn
        ? localUpcomingHolidayEmails
        : [];

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const [holidaysRes, statusRes] = await Promise.all([
          fetch("/api/holidays", { cache: "no-store" }),
          portalFetch("/api/pakistan-holiday-emails", { cache: "no-store" }),
        ]);
        const data = (await holidaysRes.json()) as HolidaysStore & {
          error?: string;
        };
        if (!holidaysRes.ok) {
          throw new Error(data.error || "Failed to load holidays.");
        }
        if (!cancelled) {
          setStore({ countries: data.countries ?? [] });
          setError(null);
        }
        if (statusRes.ok && !cancelled) {
          const statusData = (await statusRes.json()) as {
            statusByHolidayId?: Record<string, "Scheduled" | "Sent" | "—">;
            clientStatusByHolidayId?: Record<
              string,
              "Scheduled" | "Sent" | "—"
            >;
            upcoming?: PakistanHolidayEmailPreviewRow[];
          };
          setEmailStatusByHolidayId(statusData.statusByHolidayId ?? {});
          setClientEmailStatusByHolidayId(
            statusData.clientStatusByHolidayId ?? {},
          );
          setUpcomingHolidayEmails(statusData.upcoming ?? []);
        }
      } catch (loadError) {
        if (!cancelled) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load holidays.",
          );
        }
      } finally {
        if (!cancelled) {
          setReady(true);
        }
      }
    }

    void load();
    return () => {
      cancelled = true;
    };
  }, []);

  const activeCountry = useMemo(
    () =>
      store.countries.find((country) => country.id === activeCountryId) ??
      store.countries[0] ??
      null,
    [store.countries, activeCountryId],
  );

  /** Lunar rows still on auto year-shift (not yet manually edited). */
  const pendingLunarHolidaysThisYear = useMemo(() => {
    if (!activeCountry) {
      return [];
    }
    return activeCountry.holidays.filter(
      (holiday) =>
        isLunarHolidayName(holiday.name) &&
        holiday.date.startsWith(`${currentYear}-`) &&
        !isHolidayManuallySet(holiday),
    );
  }, [activeCountry, currentYear]);

  useEffect(() => {
    setLunarNoticeDismissed(false);
  }, [activeCountryId]);

  async function ensurePakistanHolidayEmails() {
    if (!canSendPakistanHolidayEmails || schedulingHolidayEmails) {
      return;
    }
    setSchedulingHolidayEmails(true);
    setInfo(null);
    try {
      // Server app-only queue + send (same path as BD cron — no portal Graph schedule).
      const response = await portalFetch("/api/pakistan-holiday-emails", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "auto-run" }),
      });
      const data = (await response.json()) as {
        statusByHolidayId?: Record<string, "Scheduled" | "Sent" | "—">;
        clientStatusByHolidayId?: Record<string, "Scheduled" | "Sent" | "—">;
        upcoming?: PakistanHolidayEmailPreviewRow[];
        message?: string;
        error?: string;
      };
      if (!response.ok) {
        throw new Error(data.error || "Failed to sync holiday emails.");
      }
      setEmailStatusByHolidayId((prev) => ({
        ...prev,
        ...(data.statusByHolidayId ?? {}),
      }));
      setClientEmailStatusByHolidayId((prev) => ({
        ...prev,
        ...(data.clientStatusByHolidayId ?? {}),
      }));
      if (data.upcoming) {
        setUpcomingHolidayEmails(data.upcoming);
      }
      if (data.message) {
        setInfo(data.message);
      }
    } catch (scheduleError) {
      setError(
        scheduleError instanceof Error
          ? scheduleError.message
          : "Failed to sync Pakistan holiday emails.",
      );
    } finally {
      setSchedulingHolidayEmails(false);
    }
  }

  useEffect(() => {
    if (
      !ready ||
      !canSendPakistanHolidayEmails ||
      activeCountryId !== "pakistan"
    ) {
      return;
    }
    void ensurePakistanHolidayEmails();
    // Run when Ayesha opens Pakistan holidays.
    // eslint-disable-next-line react-hooks/exhaustive-deps -- intentional one-shot per country visit
  }, [ready, canSendPakistanHolidayEmails, activeCountryId]);

  async function saveStore(next: HolidaysStore) {
    setSaving(true);
    setError(null);
    try {
      const saved = await persistHolidays(next);
      setStore({ countries: saved.countries ?? [] });
      setEditor(null);
    } catch (saveError) {
      setError(
        saveError instanceof Error
          ? saveError.message
          : "Failed to save holidays.",
      );
    } finally {
      setSaving(false);
    }
  }

  function updateCountryHolidays(
    countryId: HolidayCountryId,
    updater: (holidays: HolidayEntry[]) => HolidayEntry[],
  ) {
    const next: HolidaysStore = {
      countries: store.countries.map((country) =>
        country.id === countryId
          ? { ...country, holidays: updater(country.holidays) }
          : country,
      ),
    };
    void saveStore(next);
  }

  function openAddHoliday() {
    setError(null);
    setEditor({
      mode: "add",
      name: "",
      date: "",
      state: HOLIDAY_STATE_ALL,
      religion: HOLIDAY_RELIGION_ALL,
    });
  }

  function openEditHoliday(holiday: HolidayEntry) {
    setError(null);
    setEditor({
      mode: "edit",
      holidayId: holiday.id,
      name: holiday.name,
      date: holiday.date,
      state: normalizeHolidayState(holiday.state, activeCountryId),
      religion: normalizeHolidayReligion(holiday.religion),
    });
  }

  function deleteHoliday(holiday: HolidayEntry) {
    if (!activeCountry) {
      return;
    }
    if (!window.confirm(`Delete "${holiday.name}"?`)) {
      return;
    }
    updateCountryHolidays(activeCountry.id, (holidays) =>
      holidays.filter((entry) => entry.id !== holiday.id),
    );
    void logPortalAction({
      actor,
      project: "Holidays",
      entityType: "lookupData",
      entityLabel: `${activeCountry.name} / ${holiday.name}`,
      action: "deleted",
      summary: `Holiday deleted: ${holiday.name}`,
      details: `Country: ${activeCountry.name}\nHoliday: ${holiday.name}\nDate: ${formatHolidayDate(holiday.date)}`,
      comment: "Holiday removed.",
    });
  }

  function moveHoliday(holidayId: string, direction: -1 | 1) {
    if (!activeCountry) {
      return;
    }
    const index = activeCountry.holidays.findIndex(
      (holiday) => holiday.id === holidayId,
    );
    const target = index + direction;
    if (index < 0 || target < 0 || target >= activeCountry.holidays.length) {
      return;
    }
    updateCountryHolidays(activeCountry.id, (holidays) => {
      const next = [...holidays];
      const [item] = next.splice(index, 1);
      next.splice(target, 0, item);
      return next;
    });
  }

  function updateHolidayRegion(holidayId: string, region: string) {
    if (!activeCountry || !holidayUsesRegion(activeCountry.id)) {
      return;
    }
    const nextRegion = normalizeHolidayState(region, activeCountry.id);
    updateCountryHolidays(activeCountry.id, (holidays) =>
      holidays.map((holiday) =>
        holiday.id === holidayId ? { ...holiday, state: nextRegion } : holiday,
      ),
    );
    void logPortalAction({
      actor,
      project: "Holidays",
      entityType: "lookupData",
      entityLabel: `${activeCountry.name} / ${regionLabel.toLowerCase()}`,
      action: "updated",
      summary: `Holiday ${regionLabel.toLowerCase()} set to ${nextRegion}`,
      details: `Country: ${activeCountry.name}\n${regionLabel}: ${nextRegion}`,
      comment: `${activeCountry.name} holiday region scope updated.`,
    });
  }

  function updateHolidayReligion(holidayId: string, religion: string) {
    if (!activeCountry || !holidayUsesReligion(activeCountry.id)) {
      return;
    }
    const nextReligion = normalizeHolidayReligion(religion);
    updateCountryHolidays(activeCountry.id, (holidays) =>
      holidays.map((holiday) =>
        holiday.id === holidayId
          ? { ...holiday, religion: nextReligion }
          : holiday,
      ),
    );
    void logPortalAction({
      actor,
      project: "Holidays",
      entityType: "lookupData",
      entityLabel: `${activeCountry.name} / religion`,
      action: "updated",
      summary: `Holiday religion set to ${nextReligion}`,
      details: `Country: ${activeCountry.name}\nReligion: ${nextReligion}`,
      comment: "Pakistan holiday team-email religion updated.",
    });
  }

  function closeEditor() {
    if (saving) {
      return;
    }
    setEditor(null);
  }

  function submitEditor(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!editor || !activeCountry) {
      return;
    }

    const name = editor.name.trim();
    const date = normalizeHolidayDate(editor.date);
    const state = holidayUsesRegion(activeCountry.id)
      ? normalizeHolidayState(editor.state, activeCountry.id)
      : undefined;
    const religion = holidayUsesReligion(activeCountry.id)
      ? normalizeHolidayReligion(editor.religion)
      : undefined;
    if (!name) {
      setError("Name is required.");
      return;
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      setError("Enter a valid date.");
      return;
    }

    if (editor.mode === "add") {
      const holiday = createHolidayEntry(
        name,
        date,
        state,
        activeCountry.id,
        "manual",
        religion,
      );
      updateCountryHolidays(activeCountry.id, (holidays) => [
        ...holidays,
        holiday,
      ]);
      void logPortalAction({
        actor,
        project: "Holidays",
        entityType: "lookupData",
        entityLabel: `${activeCountry.name} / ${name}`,
        action: "created",
        summary: `Holiday created: ${name}`,
        details: `Country: ${activeCountry.name}\nHoliday: ${name}\nDate: ${formatHolidayDate(date)}${
          state ? `\n${regionLabel}: ${state}` : ""
        }${religion ? `\nReligion: ${religion}` : ""}`,
        comment: "New holiday added.",
      });
      return;
    }

    const previous =
      activeCountry.holidays.find((holiday) => holiday.id === editor.holidayId)
        ?.name ?? "";

    updateCountryHolidays(activeCountry.id, (holidays) =>
      holidays.map((holiday) =>
        holiday.id === editor.holidayId
          ? {
              ...holiday,
              name,
              date,
              origin: "manual" as const,
              ...(state !== undefined ? { state } : {}),
              ...(religion !== undefined ? { religion } : {}),
            }
          : holiday,
      ),
    );
    void logPortalAction({
      actor,
      project: "Holidays",
      entityType: "lookupData",
      entityLabel: `${activeCountry.name} / ${name}`,
      action: "updated",
      summary: `Holiday updated: ${name}`,
      details: `Country: ${activeCountry.name}\nName: "${previous || "—"}" → "${name}"\nDate: ${formatHolidayDate(date)}${
        state ? `\n${regionLabel}: ${state}` : ""
      }${religion ? `\nReligion: ${religion}` : ""}`,
      comment: "Holiday updated.",
    });
  }

  return (
    <>
      {!ready ? (
        <p className="table-empty">Loading…</p>
      ) : (
        <div className="lookup-layout">
          <aside className="lookup-fields" aria-label="Holiday countries">
            {HOLIDAY_COUNTRIES.map((country) => {
              const count =
                store.countries.find((entry) => entry.id === country.id)
                  ?.holidays.length ?? 0;
              return (
                <button
                  key={country.id}
                  type="button"
                  className={
                    activeCountryId === country.id
                      ? "lookup-field-link active"
                      : "lookup-field-link"
                  }
                  onClick={() => setActiveCountryId(country.id)}
                >
                  <span>{country.name}</span>
                  <span className="lookup-field-count">{count}</span>
                </button>
              );
            })}
          </aside>

          <div className="lookup-items holidays-panel">
            {activeCountry ? (
              <>
                <div className="holidays-header">
                  <div>
                    <h3 className="section-title">Holidays</h3>
                    <p className="lookup-hint">
                      {showEmailStatusColumn
                        ? `Team (Template 11): last weekday before · 9:00 AM Pakistan · Religion on each holiday (All / Muslim / Christian / Hindu) controls which User Management users get the team email. Clients (Template 12): Resource Allocation only · 7 business days before · 9:00 AM client time zone. From ${PAKISTAN_HOLIDAY_EMAIL_SENDER}.`
                        : showRegionColumn
                          ? `Federal holidays use ${regionLabel} = All. Choose a specific ${regionLabel.toLowerCase()} for local holidays.`
                          : `Add company holidays with a name and date for ${activeCountry.name}.`}
                    </p>
                  </div>
                  {canEdit ? (
                    <button
                      type="button"
                      className="button primary"
                      onClick={openAddHoliday}
                      disabled={saving}
                    >
                      + Add holiday
                    </button>
                  ) : null}
                </div>

                {error && !editor ? (
                  <p className="form-message error">{error}</p>
                ) : null}
                {info && !editor ? (
                  <p className="form-message success">{info}</p>
                ) : null}
                {showEmailStatusColumn && schedulingHolidayEmails ? (
                  <p className="lookup-hint">
                    Syncing holiday email queue with server…
                  </p>
                ) : null}

                {showEmailStatusColumn &&
                displayUpcomingHolidayEmails.length > 0 ? (
                  <div className="holiday-email-preview">
                    <h4 className="holiday-email-preview__title">
                      Next holiday email
                    </h4>
                    <p className="lookup-hint holiday-email-preview__hint">
                      Next holiday only. Team: 9:00 AM Pakistan · filtered by
                      holiday Religion (All = everyone in User Management).
                      Clients: Resource Allocation · 9:00 AM client time zone.
                      From {PAKISTAN_HOLIDAY_EMAIL_SENDER}.
                    </p>
                    <div className="table-wrap">
                      <table className="projects-table holiday-email-preview__table">
                        <thead>
                          <tr>
                            <th>Holiday</th>
                            <th>Dates</th>
                            <th>Team send</th>
                            <th>Team</th>
                            <th>Client send</th>
                            <th>Client</th>
                          </tr>
                        </thead>
                        <tbody>
                          {displayUpcomingHolidayEmails.map((row) => (
                            <tr key={row.groupKey}>
                              <td>{row.holidayName}</td>
                              <td>{row.holidayDateLabel}</td>
                              <td>{row.teamSendLabel}</td>
                              <td>
                                <span
                                  className={
                                    row.teamStatus === "Sent"
                                      ? "pakistan-holiday-email-status pakistan-holiday-email-status--sent"
                                      : row.teamStatus === "Scheduled"
                                        ? "pakistan-holiday-email-status pakistan-holiday-email-status--scheduled"
                                        : "pakistan-holiday-email-status"
                                  }
                                >
                                  {row.teamStatus}
                                </span>
                              </td>
                              <td>{row.clientSendLabel}</td>
                              <td>
                                <span
                                  className={
                                    row.clientStatus === "Sent"
                                      ? "pakistan-holiday-email-status pakistan-holiday-email-status--sent"
                                      : row.clientStatus === "Scheduled"
                                        ? "pakistan-holiday-email-status pakistan-holiday-email-status--scheduled"
                                        : "pakistan-holiday-email-status"
                                  }
                                >
                                  {row.clientStatus}
                                </span>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                ) : null}

                {pendingLunarHolidaysThisYear.length > 0 &&
                !lunarNoticeDismissed ? (
                  <div
                    className="inline-notice inline-notice--warning"
                    role="status"
                  >
                    <p className="inline-notice__text">
                      <strong>Lunar holiday:</strong> These dates need to be
                      set manually (Eid, Ashura, etc.). Edit each row below —
                      the year-shift from last year is only a placeholder.
                    </p>
                    <button
                      type="button"
                      className="inline-notice__close"
                      aria-label="Dismiss lunar holidays notice"
                      onClick={() => setLunarNoticeDismissed(true)}
                    >
                      ×
                    </button>
                  </div>
                ) : null}

                {activeCountry.holidays.length === 0 ? (
                  <p className="table-empty">
                    No holidays yet
                    {canEdit ? ". Click + Add holiday to create one." : "."}
                  </p>
                ) : (
                  <div className="table-wrap holidays-table-wrap">
                    <table className="projects-table holidays-table">
                      <thead>
                        <tr>
                          <th>Name</th>
                          <th>Date (DD-MM-YYYY)</th>
                          <th>Day</th>
                          {showRegionColumn ? <th>{regionLabel}</th> : null}
                          {showReligionColumn ? <th>Religion</th> : null}
                          {showEmailStatusColumn ? (
                            <>
                              <th>Team Email</th>
                              <th>Client Email</th>
                            </>
                          ) : null}
                          {canEdit ? <th>Action</th> : null}
                        </tr>
                      </thead>
                      <tbody>
                        {activeCountry.holidays.map((holiday, index) => {
                          const isLunar = isLunarHolidayName(holiday.name);
                          const lunarManual =
                            isLunar && isHolidayManuallySet(holiday);
                          const lunarPending = isLunar && !lunarManual;
                          return (
                          <tr
                            key={holiday.id}
                            className={
                              lunarPending
                                ? "holiday-row--lunar"
                                : lunarManual
                                  ? "holiday-row--lunar-manual"
                                  : undefined
                            }
                          >
                            <td>
                              {holiday.name}
                              {isLunar ? (
                                <span
                                  className={
                                    lunarManual
                                      ? "holidays-lunar-tag holidays-lunar-tag--manual"
                                      : "holidays-lunar-tag"
                                  }
                                >
                                  Lunar
                                </span>
                              ) : null}
                            </td>
                            <td>{formatHolidayDate(holiday.date)}</td>
                            <td>{getHolidayWeekday(holiday.date)}</td>
                            {showRegionColumn ? (
                              <td>
                                {canEdit ? (
                                  <SearchableSelect
                                    label={`${regionLabel} for ${holiday.name}`}
                                    hideLabel
                                    className="holiday-state-select"
                                    value={normalizeHolidayState(
                                      holiday.state,
                                      activeCountry.id,
                                    )}
                                    options={regionOptions}
                                    placeholder={`Search ${regionLabel.toLowerCase()}…`}
                                    disabled={saving}
                                    onChange={(value) => {
                                      if (!value) {
                                        return;
                                      }
                                      updateHolidayRegion(holiday.id, value);
                                    }}
                                  />
                                ) : (
                                  normalizeHolidayState(
                                    holiday.state,
                                    activeCountry.id,
                                  )
                                )}
                              </td>
                            ) : null}
                            {showReligionColumn ? (
                              <td>
                                {canEdit ? (
                                  <select
                                    className="field-input field-select holiday-state-select"
                                    aria-label={`Religion for ${holiday.name}`}
                                    value={normalizeHolidayReligion(
                                      holiday.religion,
                                    )}
                                    disabled={saving}
                                    onChange={(event) =>
                                      updateHolidayReligion(
                                        holiday.id,
                                        event.target.value,
                                      )
                                    }
                                  >
                                    {religionOptions.map((option) => (
                                      <option key={option} value={option}>
                                        {option}
                                      </option>
                                    ))}
                                  </select>
                                ) : (
                                  normalizeHolidayReligion(holiday.religion)
                                )}
                              </td>
                            ) : null}
                            {showEmailStatusColumn ? (
                              <>
                                <td>
                                  <span
                                    className={
                                      emailStatusByHolidayId[holiday.id] ===
                                      "Sent"
                                        ? "pakistan-holiday-email-status pakistan-holiday-email-status--sent"
                                        : emailStatusByHolidayId[holiday.id] ===
                                            "Scheduled"
                                          ? "pakistan-holiday-email-status pakistan-holiday-email-status--scheduled"
                                          : "pakistan-holiday-email-status"
                                    }
                                  >
                                    {emailStatusByHolidayId[holiday.id] ?? "—"}
                                  </span>
                                </td>
                                <td>
                                  <span
                                    className={
                                      clientEmailStatusByHolidayId[
                                        holiday.id
                                      ] === "Sent"
                                        ? "pakistan-holiday-email-status pakistan-holiday-email-status--sent"
                                        : clientEmailStatusByHolidayId[
                                              holiday.id
                                            ] === "Scheduled"
                                          ? "pakistan-holiday-email-status pakistan-holiday-email-status--scheduled"
                                          : "pakistan-holiday-email-status"
                                    }
                                  >
                                    {clientEmailStatusByHolidayId[
                                      holiday.id
                                    ] ?? "—"}
                                  </span>
                                </td>
                              </>
                            ) : null}
                            {canEdit ? (
                              <td>
                                <div className="holiday-actions">
                                  <button
                                    type="button"
                                    className="holiday-icon-btn"
                                    aria-label="Move up"
                                    title="Move up"
                                    disabled={index === 0 || saving}
                                    onClick={() => moveHoliday(holiday.id, -1)}
                                  >
                                    ↑
                                  </button>
                                  <button
                                    type="button"
                                    className="holiday-icon-btn"
                                    aria-label="Move down"
                                    title="Move down"
                                    disabled={
                                      index ===
                                        activeCountry.holidays.length - 1 ||
                                      saving
                                    }
                                    onClick={() => moveHoliday(holiday.id, 1)}
                                  >
                                    ↓
                                  </button>
                                  <button
                                    type="button"
                                    className="holiday-icon-btn"
                                    aria-label="Edit"
                                    title="Edit"
                                    disabled={saving}
                                    onClick={() => openEditHoliday(holiday)}
                                  >
                                    ✎
                                  </button>
                                  <button
                                    type="button"
                                    className="holiday-icon-btn holiday-icon-btn--danger"
                                    aria-label="Delete"
                                    title="Delete"
                                    disabled={saving}
                                    onClick={() => deleteHoliday(holiday)}
                                  >
                                    ×
                                  </button>
                                </div>
                              </td>
                            ) : null}
                          </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}
              </>
            ) : (
              <p className="table-empty">No countries configured.</p>
            )}
          </div>
        </div>
      )}

      {editor && mounted
        ? createPortal(
            <div className="dialog-backdrop" role="presentation">
              <div
                className="dialog-panel"
                role="dialog"
                aria-modal="true"
                aria-labelledby="holiday-editor-title"
              >
                <div className="dialog-header">
                  <h3 id="holiday-editor-title" className="dialog-title">
                    {editor.mode === "add" ? "Add holiday" : "Edit holiday"}
                  </h3>
                  <button
                    type="button"
                    className="dialog-close"
                    aria-label="Close"
                    onClick={closeEditor}
                    disabled={saving}
                  >
                    ×
                  </button>
                </div>

                {error ? <p className="form-message error">{error}</p> : null}

                <form
                  className="client-form client-form--dialog"
                  onSubmit={submitEditor}
                >
                  <label className="field">
                    <span className="field-label">
                      Name <span className="field-required">*</span>
                    </span>
                    <input
                      className="field-input"
                      value={editor.name}
                      onChange={(event) =>
                        setEditor({ ...editor, name: event.target.value })
                      }
                      autoFocus
                      required
                    />
                  </label>

                  <label className="field">
                    <span className="field-label">
                      Date (DD-MM-YYYY){" "}
                      <span className="field-required">*</span>
                    </span>
                    <input
                      className="field-input"
                      type="date"
                      value={editor.date}
                      onChange={(event) =>
                        setEditor({ ...editor, date: event.target.value })
                      }
                      required
                    />
                  </label>

                  {showReligionColumn ? (
                    <>
                      <label className="field">
                        <span className="field-label">Religion</span>
                        <select
                          className="field-input field-select"
                          value={editor.religion}
                          disabled={saving}
                          onChange={(event) =>
                            setEditor({
                              ...editor,
                              religion: normalizeHolidayReligion(
                                event.target.value,
                              ),
                            })
                          }
                        >
                          {religionOptions.map((option) => (
                            <option key={option} value={option}>
                              {option}
                            </option>
                          ))}
                        </select>
                      </label>
                      <p className="lookup-hint">
                        Team holiday email: All = every User Management user;
                        Muslim / Christian / Hindu = that religion only.
                      </p>
                    </>
                  ) : null}

                  {showRegionColumn ? (
                    <>
                      <SearchableSelect
                        label={regionLabel}
                        value={editor.state}
                        options={regionOptions}
                        placeholder={`Search ${regionLabel.toLowerCase()}…`}
                        required
                        disabled={saving}
                        onChange={(value) =>
                          setEditor({
                            ...editor,
                            state: value || HOLIDAY_STATE_ALL,
                          })
                        }
                      />
                      <p className="lookup-hint">
                        Use All for federal holidays that apply nationwide.
                      </p>
                    </>
                  ) : null}

                  <div className="dialog-actions">
                    <button
                      type="button"
                      className="button secondary"
                      onClick={closeEditor}
                      disabled={saving}
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      className="button primary"
                      disabled={saving}
                    >
                      {saving ? "Saving…" : "Save"}
                    </button>
                  </div>
                </form>
              </div>
            </div>,
            document.body,
          )
        : null}
    </>
  );
}
