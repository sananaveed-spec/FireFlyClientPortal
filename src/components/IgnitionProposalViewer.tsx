"use client";

import { useEffect, useState } from "react";
import { createPortal } from "react-dom";
import { portalFetch } from "@/lib/portal-api";

export type IgnitionProposalDetailView = {
  slug: string;
  name: string;
  clientName: string;
  referenceNumber: string;
  state: string;
  personalisedMessageHtml: string;
  nextStepsMessageHtml: string;
  contractTerm: string;
  minimumContractLength: number | null;
  services: {
    name: string;
    descriptionHtml: string;
    amountLabel: string;
    billingMode: string;
  }[];
};

export type IgnitionProposalViewerTarget = {
  slug: string;
  reference: string;
  clientLabel: string;
  mode: "view" | "download";
};

function safeDownloadFileName(value: string): string {
  const cleaned = value
    .trim()
    .replace(/[<>:"/\\|?*\u0000-\u001f]+/g, "-")
    .replace(/\s+/g, " ")
    .slice(0, 80);
  return cleaned || "proposal";
}

export function buildProposalDownloadHtml(
  detail: IgnitionProposalDetailView,
): string {
  const servicesHtml = detail.services
    .map((service) => {
      const amount = service.amountLabel
        ? `<p><strong>${service.amountLabel}</strong></p>`
        : "";
      const billing = service.billingMode
        ? `<p>${service.billingMode}</p>`
        : "";
      return `<section style="margin-top:1.25rem;padding-top:1rem;border-top:1px solid #ddd;">
  <h3>${service.name}</h3>
  ${amount}
  ${billing}
  ${service.descriptionHtml || ""}
</section>`;
    })
    .join("\n");

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>${detail.referenceNumber || detail.name}</title>
  <style>
    body { font-family: Georgia, "Times New Roman", serif; color: #111; max-width: 48rem; margin: 2rem auto; padding: 0 1.25rem; line-height: 1.45; }
    h1 { font-size: 1.5rem; margin-bottom: 0.5rem; }
    h2, h3, h4 { margin: 0 0 0.4rem; }
    section { break-inside: avoid; }
  </style>
</head>
<body>
  <h1>${detail.name}</h1>
  ${detail.clientName ? `<p><strong>Client:</strong> ${detail.clientName}</p>` : ""}
  ${detail.referenceNumber ? `<p><strong>Reference:</strong> ${detail.referenceNumber}</p>` : ""}
  ${detail.state ? `<p><strong>Status:</strong> ${detail.state}</p>` : ""}
  ${
    detail.minimumContractLength
      ? `<p><strong>Minimum term:</strong> ${detail.minimumContractLength} months${
          detail.contractTerm ? ` (${detail.contractTerm})` : ""
        }</p>`
      : ""
  }
  ${
    detail.personalisedMessageHtml
      ? `<section style="margin-top:1.25rem;"><h2>Message</h2>${detail.personalisedMessageHtml}</section>`
      : ""
  }
  ${servicesHtml}
  ${
    detail.nextStepsMessageHtml
      ? `<section style="margin-top:1.25rem;padding-top:1rem;border-top:1px solid #ddd;"><h2>Next steps</h2>${detail.nextStepsMessageHtml}</section>`
      : ""
  }
</body>
</html>`;
}

export function downloadProposalHtmlFile(detail: IgnitionProposalDetailView) {
  const html = buildProposalDownloadHtml(detail);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  const fileBase = safeDownloadFileName(
    detail.referenceNumber || detail.name || detail.slug,
  );
  anchor.href = url;
  anchor.download = `${fileBase}.html`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Open proposal HTML in a new window and trigger the browser print dialog. */
export function printProposalAsPdf(
  detail: IgnitionProposalDetailView,
): string | null {
  const html = buildProposalDownloadHtml(detail);
  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    return "Pop-up blocked. Allow pop-ups for this site, then try Print / Save as PDF again.";
  }
  printWindow.document.open();
  printWindow.document.write(html);
  printWindow.document.close();
  const triggerPrint = () => {
    try {
      printWindow.focus();
      printWindow.print();
    } catch {
      // Ignore; user can still print manually from the opened tab.
    }
  };
  if (printWindow.document.readyState === "complete") {
    window.setTimeout(triggerPrint, 250);
  } else {
    printWindow.addEventListener("load", () => {
      window.setTimeout(triggerPrint, 250);
    });
  }
  return null;
}

type IgnitionProposalViewerProps = {
  target: IgnitionProposalViewerTarget | null;
  onClose: () => void;
};

export function IgnitionProposalViewer({
  target,
  onClose,
}: IgnitionProposalViewerProps) {
  const [mounted, setMounted] = useState(false);
  const [detail, setDetail] = useState<IgnitionProposalDetailView | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setMounted(true);
  }, []);

  useEffect(() => {
    if (!target?.slug) {
      setDetail(null);
      setError(null);
      setLoading(false);
      return;
    }

    let cancelled = false;
    setDetail(null);
    setError(null);
    setLoading(true);

    void (async () => {
      try {
        const response = await portalFetch(
          `/api/ignition/proposals/${encodeURIComponent(target.slug)}`,
          { cache: "no-store" },
        );
        const data = (await response.json()) as {
          proposal?: IgnitionProposalDetailView;
          error?: string;
        };
        if (!response.ok || !data.proposal) {
          throw new Error(data.error || "Failed to load proposal.");
        }
        if (cancelled) {
          return;
        }
        setDetail(data.proposal);
        if (target.mode === "download") {
          downloadProposalHtmlFile(data.proposal);
        }
      } catch (loadError) {
        if (cancelled) {
          return;
        }
        setError(
          loadError instanceof Error
            ? loadError.message
            : "Failed to load proposal.",
        );
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [target]);

  if (!target || !mounted) {
    return null;
  }

  return createPortal(
    <div
      className="proposal-viewer-overlay"
      role="dialog"
      aria-modal="true"
      aria-labelledby="proposal-viewer-title"
    >
      <div className="proposal-viewer-panel">
        <div className="proposal-viewer-toolbar no-print">
          <div>
            <h3 id="proposal-viewer-title" className="section-title">
              {target.reference}
            </h3>
            <p className="lookup-hint">
              {target.clientLabel}
              {detail?.state ? ` · ${detail.state}` : ""}
            </p>
          </div>
          <div className="panel-title-actions">
            <button
              type="button"
              className="button secondary-solid"
              disabled={loading || !detail}
              onClick={() => {
                if (detail) {
                  downloadProposalHtmlFile(detail);
                }
              }}
            >
              Download
            </button>
            <button
              type="button"
              className="button secondary-solid"
              disabled={loading || !detail}
              onClick={() => {
                if (!detail) {
                  return;
                }
                const blocked = printProposalAsPdf(detail);
                if (blocked) {
                  setError(blocked);
                }
              }}
            >
              Print / Save as PDF
            </button>
            <button type="button" className="button secondary" onClick={onClose}>
              Close
            </button>
          </div>
        </div>
        {loading ? (
          <p className="table-empty">Loading proposal from Ignition…</p>
        ) : null}
        {error ? <p className="form-message error">{error}</p> : null}
        {detail ? (
          <div className="proposal-viewer-body">
            <p className="lookup-hint no-print">
              Shown inside this portal from Ignition proposal content. Download
              saves an HTML file. Print / Save as PDF opens a print dialog —
              choose &quot;Save as PDF&quot; as the printer.
            </p>
            <h2 className="proposal-viewer-heading">{detail.name}</h2>
            {detail.clientName ? (
              <p>
                <strong>Client:</strong> {detail.clientName}
              </p>
            ) : null}
            {detail.referenceNumber ? (
              <p>
                <strong>Reference:</strong> {detail.referenceNumber}
              </p>
            ) : null}
            {detail.minimumContractLength ? (
              <p>
                <strong>Minimum term:</strong> {detail.minimumContractLength}{" "}
                months
                {detail.contractTerm ? ` (${detail.contractTerm})` : ""}
              </p>
            ) : null}
            {detail.personalisedMessageHtml ? (
              <section className="proposal-viewer-section">
                <h4>Message</h4>
                <div
                  className="proposal-viewer-html"
                  dangerouslySetInnerHTML={{
                    __html: detail.personalisedMessageHtml,
                  }}
                />
              </section>
            ) : null}
            {detail.services.map((service, index) => (
              <section
                key={`${service.name}-${index}`}
                className="proposal-viewer-section"
              >
                <div className="proposal-viewer-service-head">
                  <h4>{service.name}</h4>
                  {service.amountLabel ? (
                    <strong>{service.amountLabel}</strong>
                  ) : null}
                </div>
                {service.billingMode ? (
                  <p className="lookup-hint">{service.billingMode}</p>
                ) : null}
                {service.descriptionHtml ? (
                  <div
                    className="proposal-viewer-html"
                    dangerouslySetInnerHTML={{
                      __html: service.descriptionHtml,
                    }}
                  />
                ) : null}
              </section>
            ))}
            {detail.nextStepsMessageHtml ? (
              <section className="proposal-viewer-section">
                <h4>Next steps</h4>
                <div
                  className="proposal-viewer-html"
                  dangerouslySetInnerHTML={{
                    __html: detail.nextStepsMessageHtml,
                  }}
                />
              </section>
            ) : null}
          </div>
        ) : null}
      </div>
    </div>,
    document.body,
  );
}
