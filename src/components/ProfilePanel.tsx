"use client";

import { useMsal } from "@azure/msal-react";
import { useEffect, useState } from "react";
import { useAccess } from "@/auth/accessContext";
import {
  createEmptyEmailSignature,
  DEFAULT_SIGNATURE_LOGO,
  DEFAULT_SIGNATURE_WEBSITE,
  getSignatureLogoSrc,
  getSignatureWebsiteHref,
  hasEmailSignature,
  SIGNATURE_CONFIDENTIALITY_NOTICE,
  type EmailSignature,
} from "@/lib/email-signatures";
import { portalFetch } from "@/lib/portal-api";
import { PORTAL_USERS_CHANGED_EVENT } from "@/lib/portal-users";

function SignaturePreview({ signature }: { signature: EmailSignature }) {
  const websiteHref = getSignatureWebsiteHref(signature.website);
  const logoSrc = getSignatureLogoSrc(signature);
  return (
    <div className="email-signature-block">
      {signature.name ? (
        <div className="email-signature-line email-signature-line--strong">
          {signature.name}
        </div>
      ) : null}
      {signature.email ? (
        <div className="email-signature-line email-signature-line--strong">
          {signature.email}
        </div>
      ) : null}
      {signature.position ? (
        <div className="email-signature-line email-signature-line--strong">
          {signature.position}
        </div>
      ) : null}
      {signature.phone ? (
        <div className="email-signature-line email-signature-line--strong">
          {signature.phone}
        </div>
      ) : null}
      {signature.website ? (
        <div className="email-signature-line">
          <a
            className="email-signature-link"
            href={websiteHref || DEFAULT_SIGNATURE_WEBSITE}
            target="_blank"
            rel="noreferrer"
          >
            {signature.website}
          </a>
        </div>
      ) : null}
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img
        className="email-signature-image"
        src={logoSrc}
        alt="FireFly Accountancy"
      />
      <p className="email-signature-disclaimer">
        {SIGNATURE_CONFIDENTIALITY_NOTICE}
      </p>
    </div>
  );
}

export function ProfilePanel() {
  const { accounts } = useMsal();
  const { email: accountEmail } = useAccess();
  const displayName = accounts[0]?.name ?? accountEmail;
  const [signature, setSignature] = useState<EmailSignature>(
    createEmptyEmailSignature(),
  );
  const [ready, setReady] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;

    async function load() {
      try {
        const response = await portalFetch("/api/email-signature", {
          cache: "no-store",
        });
        const data = (await response.json()) as {
          signature?: EmailSignature;
          error?: string;
        };
        if (!response.ok) {
          throw new Error(data.error || "Failed to load profile.");
        }
        if (!cancelled) {
          const loaded = data.signature ?? createEmptyEmailSignature();
          setSignature({
            ...loaded,
            name: loaded.name || displayName || "",
            email: accountEmail || loaded.email || "",
            website: DEFAULT_SIGNATURE_WEBSITE,
            imageDataUrl: DEFAULT_SIGNATURE_LOGO,
            imageName: "FireFly Accountancy logo",
          });
          setError(null);
        }
      } catch (loadError) {
        if (!cancelled) {
          setError(
            loadError instanceof Error
              ? loadError.message
              : "Failed to load profile.",
          );
        }
      } finally {
        if (!cancelled) {
          setReady(true);
        }
      }
    }

    void load();

    function onUsersChanged() {
      void load();
    }
    window.addEventListener(PORTAL_USERS_CHANGED_EVENT, onUsersChanged);

    return () => {
      cancelled = true;
      window.removeEventListener(PORTAL_USERS_CHANGED_EVENT, onUsersChanged);
    };
  }, [accountEmail, displayName]);

  return (
    <section className="content-panel content-panel--actions">
      <div className="panel-title-row">
        <div>
          <h2 className="section-title">My Profile</h2>
          <p className="lookup-hint">
            Name, position, and phone number come from User Management. Website
            and logo are fixed. This signature appears on email templates.
          </p>
        </div>
      </div>

      {!ready ? (
        <p className="table-empty">Loading…</p>
      ) : (
        <div className="profile-signature-form">
          <div className="profile-signature-card">
            <h3 className="profile-signature-heading">Email signature</h3>

            {error ? <p className="form-message error">{error}</p> : null}

            <div className="client-form-grid">
              <label className="field">
                <span className="field-label">Name</span>
                <input
                  className="field-input field-input--readonly"
                  value={signature.name}
                  readOnly
                  tabIndex={-1}
                />
              </label>

              <label className="field">
                <span className="field-label">Email</span>
                <input
                  className="field-input field-input--readonly"
                  type="email"
                  value={signature.email}
                  readOnly
                  tabIndex={-1}
                />
              </label>

              <label className="field">
                <span className="field-label">Position</span>
                <input
                  className="field-input field-input--readonly"
                  value={signature.position}
                  placeholder="Principal Operations, Consultant"
                  readOnly
                  tabIndex={-1}
                />
              </label>

              <label className="field">
                <span className="field-label">Phone Number</span>
                <input
                  className="field-input field-input--readonly"
                  value={signature.phone}
                  placeholder="(559) 333 2167"
                  readOnly
                  tabIndex={-1}
                />
              </label>

              <div className="field field--wide">
                <span className="field-label">Website</span>
                <div className="field-input field-input--readonly profile-website-readonly">
                  <a
                    className="email-signature-link"
                    href={DEFAULT_SIGNATURE_WEBSITE}
                    target="_blank"
                    rel="noreferrer"
                  >
                    {DEFAULT_SIGNATURE_WEBSITE}
                  </a>
                </div>
              </div>
            </div>

            <div className="field">
              <span className="field-label">Logo image</span>
              <div className="profile-signature-logo-auto">
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img
                  className="email-signature-image"
                  src={DEFAULT_SIGNATURE_LOGO}
                  alt="FireFly Accountancy logo"
                />
                <span className="lookup-hint">
                  FireFly Accountancy logo is applied automatically.
                </span>
              </div>
            </div>

            <div className="field field--wide">
              <span className="field-label">Confidentiality notice</span>
              <div className="field-input field-input--readonly email-signature-disclaimer-field">
                {SIGNATURE_CONFIDENTIALITY_NOTICE}
              </div>
              <span className="lookup-hint">
                This notice is fixed and appears on every email signature.
              </span>
            </div>

            {hasEmailSignature(signature) ? (
              <div className="profile-signature-preview">
                <p className="lookup-parent-label">Preview</p>
                <SignaturePreview signature={signature} />
              </div>
            ) : null}
          </div>
        </div>
      )}
    </section>
  );
}
