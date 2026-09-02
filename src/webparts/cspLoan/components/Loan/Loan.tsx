



/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-use-before-define */

import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { sp, SharePointQueryableInstance } from "@pnp/sp/presets/all";
import { InputText } from "primereact/inputtext";
import { Button } from "primereact/button";
import { Dropdown } from "primereact/dropdown";
import { Menu } from "primereact/menu";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
import { Dialog } from "primereact/dialog";
import Loader from "../Loader/Loader";
import {
  IAllDropdowns,
  IBulkTaxonomyCategory,
  IDrpdownOptions,
  IFolderDestination,
  ILoanFileMapping,
  ILoanFilterState,
  ILoanProps,
  ILoanRecord,
  ITaxonomyFieldInfo,
  ITaxonomyTag,
  ITaxonomyTagPickerProps,
  ITermStoreTerm,
  IVersionActionDialog,
  IVersionHistoryRow,
} from "../../assets/Config/interface";
import {
  listNames,
  managedMetadataFields,
  loanLibraryFields,
  parseLoanSponsorLookup,
  buildSponsorLookupUpdatePayload,
  toastFunc,
} from "../../assets/Config/Config";
import styles from "./Loan.module.scss";
import { useDispatch, useSelector } from "react-redux";
import { setLoanDetails } from "../../assets/Redux/Features/MainSPContextSlice";
import { Tree } from "primereact/tree";
import { TreeNode } from "primereact/treenode";
import { MultiSelect } from "primereact/multiselect";


// Download a blob as a file in the browser.
const triggerBlobDownload = (blob: Blob, fileName: string): void => {
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  window.setTimeout(() => window.URL.revokeObjectURL(url), 100);
};

const AVATAR_COLOR = "#4f46e5";
const LIBRARY_NAME = "testexchange";

const EMPTY_FILTER: ILoanFilterState = { search: "", sponsor: null };

const EMPTY_VERSION_ACTION: IVersionActionDialog = {
  visible: false,
  action: null,
  version: null,
};

// Parse a person display name from version/item payloads.
const parsePersonName = (field: any): string => {
  if (!field) {
    return "";
  }

  if (typeof field === "string") {
    return field.trim();
  }

  if (Array.isArray(field)) {
    return field
      .map((value) => parsePersonName(value))
      .filter(Boolean)
      .join(", ");
  }

  return (
    field.Title ||
    field.title ||
    field.Name ||
    field.name ||
    field.LookupValue ||
    field.lookupValue ||
    ""
  ).trim();
};

// Resolve a field value from item all-fields using candidate keys.
const getFieldValueFromAllFields = (
  allFields: Record<string, any>,
  ...candidates: string[]
): any => {
  for (const name of candidates) {
    if (allFields[name] !== undefined && allFields[name] !== null) {
      return allFields[name];
    }
  }

  const matchedKey = Object.keys(allFields).find((key) =>
    candidates.some(
      (candidate) => key.toLowerCase() === candidate.toLowerCase(),
    ),
  );

  return matchedKey ? allFields[matchedKey] : null;
};

// Parse SharePoint lookup strings such as "12;#Sponsor Name" or multi-value forms.
const parseLookupString = (value: string): string => {
  const trimmed = value.trim();
  if (!trimmed) {
    return "";
  }

  if (!trimmed.includes(";#")) {
    return trimmed;
  }

  const parts = trimmed.split(";#");
  const labels: string[] = [];

  for (let i = 1; i < parts.length; i += 2) {
    const label = String(parts[i] || "").trim();
    if (label) {
      labels.push(label);
    }
  }

  if (labels.length) {
    return labels.join(", ");
  }

  return parts
    .map((part) => part.trim())
    .filter((part) => part && !/^\d+$/.test(part))
    .join(", ");
};

// Parse lookup value text from different SharePoint payload shapes.
const parseLookupValue = (field: any): string => {
  if (field === null || field === undefined || field === "") {
    return "";
  }

  if (typeof field === "string") {
    return parseLookupString(field);
  }

  if (typeof field === "number") {
    return "";
  }

  if (Array.isArray(field)) {
    return field
      .map((value) => parseLookupValue(value))
      .filter(Boolean)
      .join(", ");
  }

  return (
    field.lookupValue ||
    field.LookupValue ||
    field.Title ||
    field.title ||
    field.Label ||
    field.label ||
    ""
  )
    .toString()
    .trim();
};

// Format a raw date value as dd/mm/yyyy.
const formatDateTime = (raw: any): string => {
  if (!raw) {
    return "";
  }
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
};

// Keep dates Redux-serializable (ISO string instead of Date).
const toSerializableDate = (raw: any): string | null => {
  if (!raw) {
    return null;
  }
  const date = raw instanceof Date ? raw : new Date(raw);
  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return date.toISOString();
};


// Remove trailing slash characters from a folder path.
const normalizeFolderPath = (path: string): string =>
  path.replace(/\/+$/, "");

// Decode and normalize SharePoint paths for safe comparisons.
const normalizeSharePointPath = (path: string): string => {
  if (!path) {
    return "";
  }

  try {
    return normalizeFolderPath(decodeURIComponent(path));
  } catch {
    return normalizeFolderPath(path);
  }
};

// Compare paths after SharePoint-safe normalization.
const pathsEqual = (left: string, right: string): boolean =>
  normalizeSharePointPath(left) === normalizeSharePointPath(right);

// Return the parent folder path for a given item path.
const getParentFolderPath = (itemPath: string): string => {
  const normalized = normalizeSharePointPath(itemPath);
  const lastSlash = normalized.lastIndexOf("/");
  return lastSlash > 0 ? normalized.substring(0, lastSlash) : normalized;
};

// Build destination item path from folder path and item name.
const buildDestinationItemPath = (
  folderPath: string,
  itemName: string,
): string => {
  const folder = normalizeSharePointPath(folderPath);
  const name = itemName.trim();
  return `${folder}/${name}`;
};

// Escape a text value before placing it in an OData filter.
const escapeODataString = (value: string): string => value.replace(/'/g, "''");

// Resolve SharePoint URL field payloads returned as either text or an object.
const getSharePointUrlValue = (value: any): string =>
  String(value?.Url || value?.url || value || "").trim();

// Store and open a stable absolute URL for the original SharePoint file.
const toAbsoluteSharePointUrl = (value: string): string => {
  if (!value) return "";
  try {
    return new URL(value, window.location.origin).toString();
  } catch {
    return value;
  }
};

// Open Word/Excel/PDF (and similar) in the browser viewer instead of downloading.
const toBrowserPreviewUrl = (absoluteUrl: string, fileName: string): string => {
  if (!absoluteUrl) return "";
  const extension = fileName.split(".").pop()?.toLowerCase() || "";
  const openInBrowser = [
    "doc",
    "docx",
    "xls",
    "xlsx",
    "xlsm",
    "ppt",
    "pptx",
    "pdf",
  ].includes(extension);
  if (!openInBrowser) return absoluteUrl;
  return `${absoluteUrl}${absoluteUrl.includes("?") ? "&" : "?"}web=1`;
};

const getServerRelativePathFromUrl = (value: string): string => {
  try {
    return decodeURIComponent(new URL(value, window.location.origin).pathname);
  } catch {
    return value;
  }
};

// Check whether a candidate path equals or is inside a folder path.
const isFolderDescendantOrSelf = (
  folderPath: string,
  candidatePath: string,
): boolean => {
  const parent = normalizeSharePointPath(folderPath);
  const candidate = normalizeSharePointPath(candidatePath);
  return candidate === parent || candidate.startsWith(`${parent}/`);
};

// Check whether a candidate path is within the library root path.
const isPathUnderLibrary = (
  libraryRootPath: string,
  candidatePath: string,
): boolean => {
  const root = normalizeSharePointPath(libraryRootPath);
  const candidate = normalizeSharePointPath(candidatePath);
  return candidate === root || candidate.startsWith(`${root}/`);
};

const getLoanRootPath = (
  folderPath: string,
  loanFolders: ILoanRecord[] = [],
): string | null => {
  const normalized = normalizeSharePointPath(folderPath);
  if (!normalized) {
    return null;
  }

  const matchedLoan = loanFolders
    .filter(
      (loan) =>
        !!loan.fileRef && isFolderDescendantOrSelf(loan.fileRef, normalized),
    )
    .sort(
      (left, right) =>
        normalizeSharePointPath(right.fileRef).length -
        normalizeSharePointPath(left.fileRef).length,
    )[0];

  if (matchedLoan?.fileRef) {
    return normalizeSharePointPath(matchedLoan.fileRef);
  }

  const parts = normalized.split("/").filter(Boolean);
  const libraryIndex = parts.indexOf(LIBRARY_NAME);
  if (libraryIndex < 0) {
    return null;
  }

  const loanFolderIndex = libraryIndex + 1;
  if (loanFolderIndex >= parts.length) {
    return null;
  }

  return `/${parts.slice(0, loanFolderIndex + 1).join("/")}`;
};

const itemMatchesLoanSearch = (item: ILoanRecord, keyword: string): boolean => {
  const term = keyword.trim().toLowerCase();
  if (!term) {
    return true;
  }

  return (
    item.fileName?.toLowerCase().includes(term) ||
    item.name?.toLowerCase().includes(term) ||
    item.createdby?.toLowerCase().includes(term) ||
    item.sponsor?.sponsorTitle?.toLowerCase().includes(term)
  );
};

const isItemUnderLoanRoot = (
  item: ILoanRecord,
  loanRootPath: string,
): boolean => {
  const normalizedRoot = normalizeSharePointPath(loanRootPath);
  const itemPath = normalizeSharePointPath(
    item.fileRef || item.serverRelativeUrl || "",
  );

  if (!normalizedRoot || !itemPath) {
    return false;
  }

  return isFolderDescendantOrSelf(normalizedRoot, itemPath);
};

const isOtherLoanRootFolder = (
  item: ILoanRecord,
  loanRootPath: string,
  loanFolders: ILoanRecord[],
): boolean => {
  const itemPath = normalizeSharePointPath(
    item.fileRef || item.serverRelativeUrl || "",
  );
  const normalizedRoot = normalizeSharePointPath(loanRootPath);

  return loanFolders.some(
    (loan) =>
      !!loan.fileRef &&
      loan.folderType === 1 &&
      pathsEqual(loan.fileRef, itemPath) &&
      !pathsEqual(loan.fileRef, normalizedRoot),
  );
};

const scopeItemsToLoanRoot = (
  items: ILoanRecord[],
  loanRootPath: string,
  loanFolders: ILoanRecord[] = [],
): ILoanRecord[] =>
  items.filter(
    (item) =>
      isItemUnderLoanRoot(item, loanRootPath) &&
      !isOtherLoanRootFolder(item, loanRootPath, loanFolders),
  );

const dedupeLoanRecords = (items: ILoanRecord[]): ILoanRecord[] => {
  const seen = new Set<string>();
  const result: ILoanRecord[] = [];

  items.forEach((item) => {
    const key = [
      item.id ?? "",
      normalizeSharePointPath(item.fileRef || item.serverRelativeUrl || ""),
      item.isHyperlink ? "link" : "item",
    ].join("|");

    if (seen.has(key)) {
      return;
    }

    seen.add(key);
    result.push(item);
  });

  return result;
};

const filterLoanTreeItems = (
  treeItems: ILoanRecord[],
  state: ILoanFilterState,
  loanRootPath: string,
  loanFolders: ILoanRecord[] = [],
): ILoanRecord[] => {
  let result = scopeItemsToLoanRoot(treeItems, loanRootPath, loanFolders);

  if (state.search.trim()) {
    result = result.filter((item) =>
      itemMatchesLoanSearch(item, state.search.trim()),
    );
  }

  if (state.sponsor) {
    result = result.filter(
      (item) => item.sponsor?.sponsorTitle === state.sponsor!.name,
    );
  }

  return result;
};


const SHARE_BRAND_GREEN = "#31b37c";
const SHARE_BRAND_GREEN_HOVER = "#2a9d6d";
const SHARE_BRAND_TEAL = "#1a9e98";
const SHARE_DIALOG_WIDTH = 520;
const SHARE_DIALOG_HEIGHT = 323;
const SHARE_DIALOG_BRANDING_STYLE_ID = "csp-loan-share-branding";

// Inject CSP branding styles into the native SharePoint share dialog iframe.
const injectShareDialogBranding = (iframe: HTMLIFrameElement | null): void => {
  if (!iframe?.contentDocument) {
    return;
  }

  const doc = iframe.contentDocument;
  let style = doc.getElementById(
    SHARE_DIALOG_BRANDING_STYLE_ID,
  ) as HTMLStyleElement | null;

  if (!style) {
    style = doc.createElement("style");
    style.id = SHARE_DIALOG_BRANDING_STYLE_ID;
    doc.head.appendChild(style);
  }

  style.textContent = `
    :root {
      --colorBrandBackground: ${SHARE_BRAND_GREEN} !important;
      --colorBrandBackgroundHover: ${SHARE_BRAND_GREEN_HOVER} !important;
      --colorBrandBackgroundPressed: ${SHARE_BRAND_GREEN_HOVER} !important;
      --colorCompoundBrandStroke: ${SHARE_BRAND_TEAL} !important;
      --colorCompoundBrandStrokeHover: ${SHARE_BRAND_TEAL} !important;
      --colorBrandForeground1: ${SHARE_BRAND_TEAL} !important;
    }

    html,
    body {
      margin: 0 !important;
      padding: 0 !important;
      overflow: visible !important;
      background: #ffffff !important;
    }

    .ms-Button--primary,
    button.ms-Button--primary,
    button[data-automationid="shareSendButton"],
    button[aria-label="Send"],
    button[name="Send"] {
      background-color: ${SHARE_BRAND_GREEN} !important;
      background: ${SHARE_BRAND_GREEN} !important;
      border-color: ${SHARE_BRAND_GREEN} !important;
      border-radius: 4px !important;
    }

    .ms-Button--primary .ms-Button-label,
    .ms-Button--primary .ms-Button-icon,
    .ms-Button--primary i,
    .ms-Button--primary svg,
    button[data-automationid="shareSendButton"] .ms-Button-label,
    button[data-automationid="shareSendButton"] .ms-Button-icon,
    button[data-automationid="shareSendButton"] i,
    button[data-automationid="shareSendButton"] svg,
    button[aria-label="Send"] .ms-Button-label,
    button[aria-label="Send"] .ms-Button-icon,
    button[aria-label="Send"] i,
    button[aria-label="Send"] svg {
      color: #ffffff !important;
      fill: #ffffff !important;
    }

    .ms-Button--primary:hover,
    button.ms-Button--primary:hover,
    button[data-automationid="shareSendButton"]:hover,
    button[aria-label="Send"]:hover {
      background-color: ${SHARE_BRAND_GREEN_HOVER} !important;
      background: ${SHARE_BRAND_GREEN_HOVER} !important;
      border-color: ${SHARE_BRAND_GREEN_HOVER} !important;
    }

    button[data-automationid="copyLinkButton"],
    button[aria-label*="Copy link"],
    button[aria-label*="Copy Link"] {
      background-color: #ffffff !important;
      border: 1px solid ${SHARE_BRAND_TEAL} !important;
      border-right: 1px solid ${SHARE_BRAND_TEAL} !important;
      border-radius: 4px 0 0 4px !important;
      color: ${SHARE_BRAND_TEAL} !important;
      margin-right: 0 !important;
      min-height: 32px !important;
      height: 32px !important;
      padding: 0 12px !important;
      box-shadow: none !important;
    }

    button[data-automationid="copyLinkButton"] .ms-Button-label,
    button[data-automationid="copyLinkButton"] .ms-Button-icon,
    button[data-automationid="copyLinkButton"] i,
    button[data-automationid="copyLinkButton"] svg,
    button[aria-label*="Copy link"] .ms-Button-label,
    button[aria-label*="Copy link"] .ms-Button-icon,
    button[aria-label*="Copy link"] i,
    button[aria-label*="Copy link"] svg {
      color: ${SHARE_BRAND_TEAL} !important;
      fill: ${SHARE_BRAND_TEAL} !important;
    }

    button[data-automationid="copyLinkButton"]:hover,
    button[aria-label*="Copy link"]:hover {
      background-color: rgba(26, 158, 152, 0.1) !important;
      border-color: ${SHARE_BRAND_TEAL} !important;
      color: ${SHARE_BRAND_TEAL} !important;
    }

    button[data-automationid="shareSettingsButton"],
    button[data-automationid="sharingSettingsButton"],
    button[aria-label*="Settings"],
    button[aria-label*="settings"] {
      background-color: #ffffff !important;
      border: 1px solid ${SHARE_BRAND_TEAL} !important;
      border-left: 1px solid ${SHARE_BRAND_TEAL} !important;
      border-radius: 0 4px 4px 0 !important;
      color: ${SHARE_BRAND_TEAL} !important;
      min-width: 32px !important;
      min-height: 32px !important;
      height: 32px !important;
      margin-left: -1px !important;
      padding: 0 8px !important;
      box-shadow: none !important;
    }

    button[data-automationid="shareSettingsButton"] .ms-Button-icon,
    button[data-automationid="shareSettingsButton"] i,
    button[data-automationid="shareSettingsButton"] svg,
    button[aria-label*="Settings"] .ms-Button-icon,
    button[aria-label*="Settings"] i,
    button[aria-label*="Settings"] svg {
      color: ${SHARE_BRAND_TEAL} !important;
      fill: ${SHARE_BRAND_TEAL} !important;
    }

    button[data-automationid="shareSettingsButton"]:hover,
    button[aria-label*="Settings"]:hover {
      background-color: rgba(26, 158, 152, 0.1) !important;
      border-color: ${SHARE_BRAND_TEAL} !important;
      color: ${SHARE_BRAND_TEAL} !important;
    }

    .ms-TextField-fieldGroup:focus-within,
    .ms-TextField.is-active .ms-TextField-fieldGroup,
    .ms-TextField-fieldGroup.ms-TextField-fieldGroup--focused {
      border-color: ${SHARE_BRAND_TEAL} !important;
    }

    .ms-TextField-fieldGroup:focus-within::after {
      border-color: ${SHARE_BRAND_TEAL} !important;
    }

    .ms-Link,
    a.ms-Link {
      color: ${SHARE_BRAND_TEAL} !important;
    }

    button[data-automationid="closeButton"],
    button[data-automationid="helpButton"],
    button[data-automationid="overflowButton"],
    button.ms-Dialog-button--close {
      color: #64748b !important;
      background: transparent !important;
      border: none !important;
      border-radius: 4px !important;
    }

    button[data-automationid="closeButton"]:hover,
    button[data-automationid="helpButton"]:hover,
    button[data-automationid="overflowButton"]:hover,
    button.ms-Dialog-button--close:hover {
      color: ${SHARE_BRAND_TEAL} !important;
      background: rgba(26, 158, 152, 0.08) !important;
    }

    [role="heading"],
    h1,
    h2 {
      color: #111827 !important;
    }
  `;
};

// Apply runtime button branding overrides inside the share dialog iframe.
const brandShareDialogButtons = (iframe: HTMLIFrameElement | null): void => {
  const doc = iframe?.contentDocument;
  if (!doc) {
    return;
  }

  const paintChildren = (btn: HTMLElement, color: string, fillWhite = false): void => {
    const tone = fillWhite ? "#ffffff" : color;
    btn.querySelectorAll(".ms-Button-label, .ms-Button-icon, i, svg, span").forEach(
      (node) => {
        const el = node as HTMLElement;
        el.style.setProperty("color", tone, "important");
        el.style.setProperty("fill", tone, "important");
      },
    );
  };

  doc.querySelectorAll("button").forEach((node) => {
    const btn = node as HTMLButtonElement;
    const label = (
      btn.getAttribute("aria-label") ||
      btn.getAttribute("data-automationid") ||
      btn.innerText ||
      ""
    ).toLowerCase();

    if (
      label.includes("send") ||
      btn.classList.contains("ms-Button--primary") ||
      btn.getAttribute("data-automationid") === "shareSendButton"
    ) {
      btn.style.setProperty("background-color", SHARE_BRAND_GREEN, "important");
      btn.style.setProperty("background", SHARE_BRAND_GREEN, "important");
      btn.style.setProperty("border-color", SHARE_BRAND_GREEN, "important");
      paintChildren(btn, SHARE_BRAND_GREEN, true);
      return;
    }

    if (label.includes("copy link") || label.includes("copylink")) {
      btn.style.setProperty("background-color", "#ffffff", "important");
      btn.style.setProperty("border", `1px solid ${SHARE_BRAND_TEAL}`, "important");
      btn.style.setProperty("border-radius", "4px 0 0 4px", "important");
      btn.style.setProperty("color", SHARE_BRAND_TEAL, "important");
      paintChildren(btn, SHARE_BRAND_TEAL);
      return;
    }

    if (
      label.includes("setting") ||
      btn.getAttribute("data-automationid")?.toLowerCase().includes("setting") ||
      !!btn.querySelector('[data-icon-name="Settings"], [data-icon-name="Gear"]')
    ) {
      btn.style.setProperty("background-color", "#ffffff", "important");
      btn.style.setProperty("border", `1px solid ${SHARE_BRAND_TEAL}`, "important");
      btn.style.setProperty("border-radius", "0 4px 4px 0", "important");
      btn.style.setProperty("color", SHARE_BRAND_TEAL, "important");
      paintChildren(btn, SHARE_BRAND_TEAL);
    }
  });
};

// Resize share dialog iframe to the expected modal dimensions.
const fitShareDialogIframe = (iframe: HTMLIFrameElement | null): void => {
  if (iframe) {
    iframe.style.width = `${SHARE_DIALOG_WIDTH}px`;
    iframe.style.height = `${SHARE_DIALOG_HEIGHT}px`;
  }
};

// Schedule multiple delayed branding passes for async share dialog content.
const scheduleShareDialogBranding = (
  iframe: HTMLIFrameElement | null,
): void => {
  const apply = (): void => {
    injectShareDialogBranding(iframe);
    brandShareDialogButtons(iframe);
    fitShareDialogIframe(iframe);
  };

  apply();
  [300, 800, 1500, 2500].forEach((delay) => {
    window.setTimeout(apply, delay);
  });
};



const GUID_REGEX =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// True when SharePoint returned a WssId instead of a readable taxonomy label.
const isNumericTaxonomyLabel = (label?: string): boolean =>
  Boolean(label && /^\d+$/.test(label.trim()));

// Parse managed metadata / taxonomy wire formats into readable labels.
const parseTaxonomyLabels = (field: any): string[] => {
  if (field === null || field === undefined || field === "") {
    return [];
  }

  if (typeof field === "number") {
    return [];
  }

  if (typeof field === "string") {
    const trimmed = field.trim();
    if (!trimmed) {
      return [];
    }

    // Wire formats:
    // "Label|guid"
    // "-1;#Label|guid"
    // "5;#Label|guid;6;#Label2|guid2"
    // "Label1; Label2" (FieldValuesAsText)
    if (trimmed.includes("|") || trimmed.includes(";#")) {
      const chunks = trimmed.split(";#");
      const labels: string[] = [];

      chunks.forEach((chunk) => {
        const piece = chunk.trim();
        if (!piece || piece === "-1" || /^\d+$/.test(piece)) {
          return;
        }

        const labelPart = piece.includes("|")
          ? piece.split("|")[0].trim()
          : piece;
        if (
          labelPart &&
          !isNumericTaxonomyLabel(labelPart) &&
          !GUID_REGEX.test(labelPart)
        ) {
          labels.push(labelPart);
        }
      });

      if (labels.length) {
        return labels;
      }
    }

    return trimmed
      .split(";")
      .map((value) => value.trim())
      .filter(
        (value) =>
          value && !isNumericTaxonomyLabel(value) && !GUID_REGEX.test(value),
      );
  }

  if (Array.isArray(field)) {
    return field
      .reduce<string[]>(
        (labels, value) => labels.concat(parseTaxonomyLabels(value)),
        [],
      )
      .filter(Boolean);
  }

  if (Array.isArray(field.results)) {
    return parseTaxonomyLabels(field.results);
  }

  const label = String(field.Label || field.label || field.lookupValue || "")
    .trim();
  if (label && !isNumericTaxonomyLabel(label) && !GUID_REGEX.test(label)) {
    return [label];
  }

  return [];
};

// Convert taxonomy labels into a comma-separated display string.
const parseTaxonomyValue = (field: any): string => {
  return parseTaxonomyLabels(field).join(", ");
};

// Collect taxonomy WssIds from a field payload for TaxonomyHiddenList lookup.
const collectTaxonomyWssIds = (field: any): number[] => {
  if (field === null || field === undefined || field === "") {
    return [];
  }

  if (typeof field === "number" && field > 0) {
    return [field];
  }

  if (typeof field === "string") {
    const ids: number[] = [];
    const parts = field.split(";#");
    parts.forEach((part) => {
      const trimmed = part.trim();
      if (/^\d+$/.test(trimmed) && trimmed !== "-1") {
        ids.push(Number(trimmed));
      }
    });
    return ids;
  }

  if (Array.isArray(field)) {
    return field.reduce<number[]>(
      (ids, value) => ids.concat(collectTaxonomyWssIds(value)),
      [],
    );
  }

  // OData verbose collection shape.
  if (Array.isArray(field.results)) {
    return collectTaxonomyWssIds(field.results);
  }

  const wssId = Number(field.WssId ?? field.wssId ?? 0);
  if (wssId > 0) {
    return [wssId];
  }

  if (isNumericTaxonomyLabel(String(field.Label || field.label || ""))) {
    return [Number(field.Label || field.label)];
  }

  return [];
};

// Collect taxonomy term GUIDs from a field payload.
const collectTaxonomyTermGuids = (field: any): string[] => {
  if (field === null || field === undefined || field === "") {
    return [];
  }

  if (typeof field === "string") {
    const guids: string[] = [];
    const guidMatches = field.match(
      /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi,
    );
    if (guidMatches?.length) {
      guids.push(...guidMatches.map((guid) => guid.toLowerCase()));
    }
    return guids;
  }

  if (Array.isArray(field)) {
    return field.reduce<string[]>(
      (guids, value) => guids.concat(collectTaxonomyTermGuids(value)),
      [],
    );
  }

  if (Array.isArray(field.results)) {
    return collectTaxonomyTermGuids(field.results);
  }

  const termGuid = String(field.TermGuid || field.termGuid || "")
    .replace(/[{}]/g, "")
    .trim()
    .toLowerCase();
  return GUID_REGEX.test(termGuid) ? [termGuid] : [];
};

// Find taxonomy-related raw values on a version payload (handles alternate keys).
const getTaxonomyFieldCandidates = (
  payload: Record<string, any> | null | undefined,
  fieldInternalName: string,
): any[] => {
  if (!payload) {
    return [];
  }

  const noteFieldName = `${fieldInternalName}_0`;
  const keys = Object.keys(payload);
  const normalizedTarget = fieldInternalName.toLowerCase();
  const matchedKeys = keys.filter((key) => {
    const normalizedKey = key.toLowerCase().replace(/\.$/, "");
    return (
      normalizedKey === normalizedTarget ||
      normalizedKey === `${normalizedTarget}_0` ||
      normalizedKey === noteFieldName.toLowerCase() ||
      normalizedKey.startsWith(`${normalizedTarget}_`) ||
      // Catch OData-encoded variants for "Asset management"
      (normalizedTarget.includes("asset") &&
        normalizedKey.includes("asset") &&
        normalizedKey.includes("management"))
    );
  });

  return [
    payload[fieldInternalName],
    payload[noteFieldName],
    ...matchedKeys.map((key) => payload[key]),
  ].filter((value) => value !== undefined && value !== null && value !== "");
};

// Resolve Sponsor lookup text from a version/item payload.
const resolveSponsorFromVersion = (
  versionItem: Record<string, any>,
  textValues?: Record<string, any> | null,
  sponsorIdTitleMap?: Map<number, string>,
): string => {
  const fromText =
    parseLookupValue(textValues?.[loanLibraryFields.sponsorName]) ||
    parseLookupValue(textValues?.sponsor);
  if (fromText) {
    return fromText;
  }

  const fromField =
    parseLookupValue(versionItem[loanLibraryFields.sponsorName]) ||
    parseLookupValue(
      getFieldValueFromAllFields(versionItem, loanLibraryFields.sponsorName),
    );
  if (fromField) {
    return fromField;
  }

  const sponsorId = Number(
    versionItem[loanLibraryFields.sponsorNameId] ??
      versionItem[loanLibraryFields.sponsorName]?.LookupId ??
      versionItem[loanLibraryFields.sponsorName]?.lookupId ??
      textValues?.[loanLibraryFields.sponsorNameId] ??
      0,
  );

  if (sponsorId > 0 && sponsorIdTitleMap?.has(sponsorId)) {
    return sponsorIdTitleMap.get(sponsorId) || "";
  }

  return "";
};

// Resolve managed metadata labels from version/item payload (+ optional maps).
const resolveTaxonomyFromVersion = (
  versionItem: Record<string, any>,
  fieldInternalName: string,
  textValues?: Record<string, any> | null,
  wssIdLabelMap?: Map<number, string>,
  termGuidLabelMap?: Map<string, string>,
): string => {
  const noteFieldName = `${fieldInternalName}_0`;
  const displayNameGuess = fieldInternalName
    .replace(/_x0020_/gi, " ")
    .replace(/_/g, " ")
    .trim();

  const candidates = [
    ...getTaxonomyFieldCandidates(textValues, fieldInternalName),
    textValues?.[fieldInternalName],
    textValues?.[noteFieldName],
    textValues?.[displayNameGuess],
    ...getTaxonomyFieldCandidates(versionItem, fieldInternalName),
  ];

  // Prefer readable text values first.
  for (const candidate of candidates) {
    const parsed = parseTaxonomyValue(candidate);
    if (parsed) {
      return parsed;
    }
  }

  // Resolve by WssId (common on historical versions).
  if (wssIdLabelMap?.size) {
    const ids = candidates.reduce<number[]>(
      (all, candidate) => all.concat(collectTaxonomyWssIds(candidate)),
      [],
    );
    const labels = Array.from(new Set(ids))
      .map((id) => wssIdLabelMap.get(id) || "")
      .filter(Boolean);
    if (labels.length) {
      return labels.join(", ");
    }
  }

  // Resolve by TermGuid when Label is missing/numeric.
  if (termGuidLabelMap?.size) {
    const guids = candidates.reduce<string[]>(
      (all, candidate) => all.concat(collectTaxonomyTermGuids(candidate)),
      [],
    );
    const labels = Array.from(new Set(guids))
      .map((guid) => termGuidLabelMap.get(guid) || "")
      .filter(Boolean);
    if (labels.length) {
      return labels.join(", ");
    }
  }

  return "";
};

// Load TaxonomyHiddenList titles for WssIds and/or term GUIDs.
const resolveTaxonomyLabelMaps = async (
  wssIds: number[],
  termGuids: string[],
): Promise<{
  wssIdLabelMap: Map<number, string>;
  termGuidLabelMap: Map<string, string>;
}> => {
  const uniqueIds = Array.from(new Set(wssIds.filter((id) => id > 0)));
  const uniqueGuids = Array.from(
    new Set(
      termGuids
        .map((guid) => guid.replace(/[{}]/g, "").trim().toLowerCase())
        .filter((guid) => GUID_REGEX.test(guid)),
    ),
  );

  const wssIdLabelMap = new Map<number, string>();
  const termGuidLabelMap = new Map<string, string>();

  if (!uniqueIds.length && !uniqueGuids.length) {
    return { wssIdLabelMap, termGuidLabelMap };
  }

  const applyRows = (rows: any[]): void => {
    (rows || []).forEach((row: any) => {
      const id = Number(row.Id ?? row.ID);
      const title = String(row.Title || "").trim();
      const termId = String(row.IdForTerm || "")
        .replace(/[{}]/g, "")
        .trim()
        .toLowerCase();

      if (id > 0 && title) {
        wssIdLabelMap.set(id, title);
      }
      if (termId && title) {
        termGuidLabelMap.set(termId, title);
      }
    });
  };

  const tryLoad = async (list: any): Promise<boolean> => {
    try {
      const filters: string[] = [];
      if (uniqueIds.length) {
        filters.push(uniqueIds.map((id) => `Id eq ${id}`).join(" or "));
      }
      if (uniqueGuids.length) {
        filters.push(
          uniqueGuids.map((guid) => `IdForTerm eq '${guid}'`).join(" or "),
        );
      }

      const filter = filters.map((part) => `(${part})`).join(" or ");
      const rows: any[] = await list.items
        .select("Id", "Title", "IdForTerm")
        .filter(filter)
        .top(Math.max(uniqueIds.length + uniqueGuids.length, 1))();

      applyRows(rows);
      return true;
    } catch (error) {
      console.warn("TaxonomyHiddenList lookup failed:", error);
      return false;
    }
  };

  const rootLoaded = await tryLoad(sp.site.rootWeb.lists.getByTitle("TaxonomyHiddenList"));
  if (!rootLoaded || (!wssIdLabelMap.size && !termGuidLabelMap.size)) {
    await tryLoad(sp.web.lists.getByTitle("TaxonomyHiddenList"));
  }

  return { wssIdLabelMap, termGuidLabelMap };
};

// Read a specific list-item version payload (includes lookup/taxonomy fields).
const getListItemVersionPayload = async (
  itemId: number,
  itemVersionId: number,
): Promise<Record<string, any> | null> => {
  if (!itemId || !itemVersionId) {
    return null;
  }

  try {
    return await sp.web.lists
      .getByTitle(listNames.loan)
      .items.getById(itemId)
      .versions.getById(itemVersionId)();
  } catch (error) {
    console.warn("List item version lookup failed:", error);
    return null;
  }
};

// Read FieldValuesAsText for a specific list-item version.
const getListItemVersionFieldValuesAsText = async (
  itemId: number,
  itemVersionId: number,
): Promise<Record<string, any> | null> => {
  if (!itemId || !itemVersionId) {
    return null;
  }

  try {
    const versionQueryable = sp.web.lists
      .getByTitle(listNames.loan)
      .items.getById(itemId)
      .versions.getById(itemVersionId);

    return await SharePointQueryableInstance(
      versionQueryable,
      "FieldValuesAsText",
    )();
  } catch (error) {
    console.warn("Version FieldValuesAsText lookup failed:", error);
    return null;
  }
};

// Split a comma-separated taxonomy string into individual labels.
const toTaxonomyLabels = (value?: string): string[] => {
  if (!value) return [];
  return value
    .split(",")
    .map((label) => label.trim())
    .filter(Boolean);
};

// Convert a SharePoint version label (for example 2.0 or 2.3) to its
// immutable _vti_history version identifier.
const getFileVersionHistoryId = (versionLabel: string): number => {
  const [majorText, minorText = "0"] = versionLabel.split(".");
  const majorVersion = Number(majorText);
  const minorVersion = Number(minorText);

  if (!Number.isFinite(majorVersion) || !Number.isFinite(minorVersion)) {
    return 0;
  }

  return majorVersion * 512 + minorVersion;
};

// Build the server-relative URL for an exact historical file version.
const buildHistoricalFileUrl = (
  fileServerRelativeUrl: string,
  webServerRelativeUrl: string,
  versionLabel: string,
): string => {
  const versionId = getFileVersionHistoryId(versionLabel);
  if (!versionId || !fileServerRelativeUrl) return "";

  const normalizedWebUrl =
    webServerRelativeUrl === "/"
      ? ""
      : webServerRelativeUrl.replace(/\/+$/, "");
  const filePathWithinWeb = fileServerRelativeUrl.startsWith(normalizedWebUrl)
    ? fileServerRelativeUrl.slice(normalizedWebUrl.length)
    : fileServerRelativeUrl;

  return `${normalizedWebUrl}/_vti_history/${versionId}${
    filePathWithinWeb.startsWith("/") ? filePathWithinWeb : `/${filePathWithinWeb}`
  }`;
};

const isOfficeDocumentName = (fileName: string): boolean => {
  const extension = fileName.split(".").pop()?.toLowerCase() || "";
  return ["doc", "docx", "xls", "xlsx", "xlsm", "ppt", "pptx"].includes(
    extension,
  );
};

const getOfficeDesktopProtocol = (fileName: string): string => {
  const extension = fileName.split(".").pop()?.toLowerCase() || "";
  if (["doc", "docx"].includes(extension)) return "ms-word";
  if (["xls", "xlsx", "xlsm"].includes(extension)) return "ms-excel";
  if (["ppt", "pptx"].includes(extension)) return "ms-powerpoint";
  return "";
};



// Shared CAML view fields used by root and subfolder queries.
const VIEW_FIELDS_XML = `
  <FieldRef Name="ID" /><FieldRef Name="FileRef" /><FieldRef Name="FileLeafRef" />
  <FieldRef Name="FSObjType" /><FieldRef Name="Author" /><FieldRef Name="Created" />
  <FieldRef Name="Modified" /><FieldRef Name="${loanLibraryFields.sponsorName}" />
  <FieldRef Name="${managedMetadataFields.assetManagement}" />
  <FieldRef Name="${managedMetadataFields.legal}" />
  <FieldRef Name="${managedMetadataFields.servicing}" />
`;

// Resolve taxonomy text from a renderListDataAsStream row.
const getTaxonomyFromStreamRow = (item: any, fieldName: string): string => {
  const itemKeys = Object.keys(item || {});
  const matchedKey = itemKeys.find(
    (key) =>
      key.replace(/\.$/, "").toLowerCase() === fieldName.toLowerCase(),
  );

  const candidates = [
    item[fieldName],
    item[`${fieldName}.`],
    matchedKey ? item[matchedKey] : undefined,
    item[`ows_${fieldName}`],
    item[`OData_${fieldName}`],
  ];

  for (const value of candidates) {
    const parsed = parseTaxonomyValue(value);
    if (parsed) {
      return parsed;
    }
  }

  return "";
};

// Map a raw renderListDataAsStream row into ILoanRecord.
const mapRowToLoanRecord = (item: any): ILoanRecord => {
  const sponsorLookup = parseLoanSponsorLookup(item);

  return {
    id: Number(item.ID),
    name: item.FileLeafRef || "",
    fileName: item.FileLeafRef || "",
    fileRef: item.FileRef || "",
    serverRelativeUrl: item.FileRef || "",
    folderType: Number(item.FSObjType) || 0,
    sponsor: {
      id: sponsorLookup.id || null,
      sponsorTitle: sponsorLookup.sponsorTitle,
    },
    createdby: item.Author?.[0]?.title || item.Author?.title || "",
    createddate: toSerializableDate(item.Created),
    modifieddate: toSerializableDate(item.Modified),
    assetmanagement: getTaxonomyFromStreamRow(
      item,
      managedMetadataFields.assetManagement,
    ),
    servicing: getTaxonomyFromStreamRow(item, managedMetadataFields.servicing),
    legal: getTaxonomyFromStreamRow(item, managedMetadataFields.legal),
    type: "existing",
  };
};



// ─── File upload managed metadata helpers ───────────────────────────────────
type TaxonomyCategory = IBulkTaxonomyCategory;

const METADATA_TAG_CATEGORIES = {
  "Asset Management": [managedMetadataFields.assetManagement],
  Legal: [managedMetadataFields.legal],
  Servicing: [managedMetadataFields.servicing],
} as const;

const TAXONOMY_CATEGORIES = Object.keys(
  METADATA_TAG_CATEGORIES,
) as TaxonomyCategory[];

const EMPTY_EDIT_TAG_TERM_IDS: Record<TaxonomyCategory, string[]> = {
  "Asset Management": [],
  Legal: [],
  Servicing: [],
};

const TAXONOMY_CATEGORY_ROW_FIELD: Record<
  TaxonomyCategory,
  "assetmanagement" | "legal" | "servicing"
> = {
  "Asset Management": "assetmanagement",
  Legal: "legal",
  Servicing: "servicing",
};

const normalizeTaxonomyGuid = (guid: string | undefined): string =>
  (guid || "").replace(/[{}]/g, "").trim().toLowerCase();

const termIdsMatch = (left: string, right: string): boolean =>
  left.trim().toLowerCase() === right.trim().toLowerCase();

const getTermId = (term: ITermStoreTerm): string =>
  normalizeTaxonomyGuid(term.id || term.Id);

const getDefaultTermLabel = (term: ITermStoreTerm): string => {
  const labels = term.labels || [];
  const preferred =
    labels.find((label) => label.isDefault)?.name || labels[0]?.name || "";
  return preferred.trim();
};

const mapTermToTreeNode = (term: ITermStoreTerm): TreeNode => {
  const termId = getTermId(term);
  const children = (term.children || []).map(mapTermToTreeNode);
  const label = getDefaultTermLabel(term);
  return {
    key: termId,
    label,
    data: { termId, label },
    children,
    leaf: children.length === 0,
    selectable: true,
  };
};

const buildTreeFromFlatTerms = (terms: ITermStoreTerm[]): TreeNode[] => {
  const termMap = new Map<string, TreeNode>();
  const roots: TreeNode[] = [];

  for (const term of terms) {
    const termId = getTermId(term);
    if (!termId) continue;
    termMap.set(termId, mapTermToTreeNode({ ...term, children: [] }));
  }

  for (const term of terms) {
    const termId = getTermId(term);
    if (!termId || !termMap.has(termId)) continue;
    const node = termMap.get(termId)!;
    const parentId = normalizeTaxonomyGuid(term.parent?.id);
    if (parentId && termMap.has(parentId)) {
      const parent = termMap.get(parentId)!;
      parent.children = parent.children || [];
      parent.children.push(node);
      parent.leaf = false;
    } else {
      roots.push(node);
    }
  }

  return roots;
};

const fetchTermStoreJson = async (url: string): Promise<any | null> => {
  const response = await fetch(url, {
    method: "GET",
    headers: { Accept: "application/json;odata=nometadata" },
    credentials: "same-origin",
  });
  if (!response.ok) return null;
  return response.json();
};

const loadLegacyTermChildren = async (
  siteUrl: string,
  termSetId: string,
  parentTermId?: string,
): Promise<ITermStoreTerm[]> => {
  const setId = normalizeTaxonomyGuid(termSetId);
  const parentPath = parentTermId
    ? `/terms/${normalizeTaxonomyGuid(parentTermId)}/getlegacychildren`
    : "/getlegacychildren";
  const payload = await fetchTermStoreJson(
    `${siteUrl}/_api/v2.1/termStore/termSets/${setId}${parentPath}?$select=id,labels,childrenCount`,
  );
  return payload?.value || [];
};

const loadTermSetTreeLegacy = async (
  siteUrl: string,
  termSetId: string,
  parentTermId?: string,
): Promise<TreeNode[]> => {
  const terms = await loadLegacyTermChildren(siteUrl, termSetId, parentTermId);
  const nodes: TreeNode[] = [];

  for (const term of terms) {
    const termId = getTermId(term);
    if (!termId) continue;
    const children =
      (term.childrenCount || 0) > 0
        ? await loadTermSetTreeLegacy(siteUrl, termSetId, termId)
        : [];
    const label = getDefaultTermLabel(term);
    nodes.push({
      key: termId,
      label,
      data: { termId, label },
      children,
      leaf: children.length === 0,
      selectable: true,
    });
  }

  return nodes;
};

const loadTermSetTree = async (
  termSetId: string,
  groupId?: string,
): Promise<TreeNode[]> => {
  const { Url } = await sp.web.select("Url")();
  const setId = normalizeTaxonomyGuid(termSetId);
  let group = groupId ? normalizeTaxonomyGuid(groupId) : "";

  if (!group) {
    const setMeta = await fetchTermStoreJson(
      `${Url}/_api/v2.1/termStore/sets/${setId}`,
    );
    group = normalizeTaxonomyGuid(setMeta?.parent?.id || setMeta?.groupId);
  }

  const urls = [
    group
      ? `${Url}/_api/v2.1/termStore/groups/${group}/sets/${setId}/terms?$expand=children($levels=max)`
      : "",
    `${Url}/_api/v2.1/termStore/sets/${setId}/terms?$expand=children($levels=max)`,
    `${Url}/_api/v2.1/termStore/sets/${setId}/terms?$top=500`,
  ].filter(Boolean);

  for (const url of urls) {
    const payload = await fetchTermStoreJson(url);
    const terms: ITermStoreTerm[] = payload?.value || [];
    if (!terms.length) continue;

    const hasNestedChildren = terms.some(
      (term) => (term.children || []).length > 0,
    );
    if (hasNestedChildren) {
      return terms.map(mapTermToTreeNode);
    }

    if (terms.some((term) => term.parent?.id)) {
      const flatTree = buildTreeFromFlatTerms(terms);
      if (flatTree.length) return flatTree;
    }

    return terms.map(mapTermToTreeNode);
  }

  return loadTermSetTreeLegacy(Url, setId);
};

const resolveTaxonomyFieldForCategory = async (
  category: TaxonomyCategory,
): Promise<ITaxonomyFieldInfo | null> => {
  for (const name of METADATA_TAG_CATEGORIES[category]) {
    try {
      const field: any = await sp.web.lists
        .getByTitle(listNames.loan)
        .fields.getByInternalNameOrTitle(name)
        .select("InternalName", "TermSetId", "SspId", "TypeAsString")();

      if (field?.TermSetId) {
        return {
          category,
          internalName: field.InternalName,
          termSetId: normalizeTaxonomyGuid(field.TermSetId),
          groupId: field.SspId
            ? normalizeTaxonomyGuid(field.SspId)
            : undefined,
        };
      }
    } catch {
      // Try next candidate.
    }
  }
  return null;
};

const loadManagedMetadataTagTrees = async (): Promise<{
  trees: Record<TaxonomyCategory, TreeNode[]>;
  fields: Record<TaxonomyCategory, string>;
}> => {
  const trees = {} as Record<TaxonomyCategory, TreeNode[]>;
  const fields = {} as Record<TaxonomyCategory, string>;
  const fieldInfos = (
    await Promise.all(
      TAXONOMY_CATEGORIES.map((category) =>
        resolveTaxonomyFieldForCategory(category),
      ),
    )
  ).filter((field): field is ITaxonomyFieldInfo => field !== null);

  await Promise.all(
    fieldInfos.map(async (info) => {
      try {
        trees[info.category] = await loadTermSetTree(
          info.termSetId,
          info.groupId,
        );
        fields[info.category] = info.internalName;
      } catch (error) {
        console.error(`loadManagedMetadataTagTrees ${info.category}:`, error);
        trees[info.category] = [];
      }
    }),
  );

  return { trees, fields };
};

let cachedTreesPromise: Promise<{
  trees: Record<TaxonomyCategory, TreeNode[]>;
  fields: Record<TaxonomyCategory, string>;
}> | null = null;

const loadManagedMetadataTagTreesCached = (): Promise<{
  trees: Record<TaxonomyCategory, TreeNode[]>;
  fields: Record<TaxonomyCategory, string>;
}> => {
  if (!cachedTreesPromise) {
    cachedTreesPromise = loadManagedMetadataTagTrees();
  }
  return cachedTreesPromise;
};

const findTermInTree = (
  nodes: TreeNode[],
  termId: string,
): TreeNode | null => {
  for (const node of nodes) {
    if (termIdsMatch(String(node.key), termId)) return node;
    if (node.children?.length) {
      const match = findTermInTree(node.children, termId);
      if (match) return match;
    }
  }
  return null;
};

const findNodeLabel = (nodes: TreeNode[], termId: string): string => {
  for (const node of nodes) {
    if (termIdsMatch(String(node.key), termId)) {
      return String(node.label || node.data?.label || "");
    }
    if (node.children?.length) {
      const childLabel = findNodeLabel(node.children, termId);
      if (childLabel) return childLabel;
    }
  }
  return "";
};

const buildTaxonomyTags = (
  termIds: string[],
  category: TaxonomyCategory,
  trees: Record<TaxonomyCategory, TreeNode[]>,
  fields: Record<TaxonomyCategory, string>,
): ITaxonomyTag[] =>
  termIds
    .map((termId) => {
      const node = findTermInTree(trees[category] || [], termId);
      if (node?.data && fields[category]) {
        return {
          termId: String(node.key),
          label: node.data.label || String(node.label),
          fieldInternalName: fields[category],
          category,
        };
      }
      if (!fields[category]) return null;
      return {
        termId,
        label: findNodeLabel(trees[category] || [], termId) || termId,
        fieldInternalName: fields[category],
        category,
      };
    })
    .filter((tag): tag is ITaxonomyTag => tag !== null);

const formatTaxonomyFieldValue = (label: string, termId: string): string => {
  const termGuid = termId.replace(/[{}]/g, "").trim();
  return `${label}|${termGuid}`;
};

const getValidateUpdateFailures = (result: any): any[] => {
  const entries = Array.isArray(result?.value)
    ? result.value
    : Array.isArray(result?.d?.ValidateUpdateListItem?.results)
      ? result.d.ValidateUpdateListItem.results
      : Array.isArray(result)
        ? result
        : [];
  return entries.filter(
    (entry: any) => entry?.HasException || entry?.ErrorMessage,
  );
};

const applyTaxonomyToListItem = async (
  itemId: number,
  tags: ITaxonomyTag[],
): Promise<void> => {
  if (!tags.length) return;
  const fieldInternalName = tags[0].fieldInternalName;
  if (!fieldInternalName) {
    throw new Error(`No SharePoint field mapped for ${tags[0].category}`);
  }

  const wireValues = tags.map((tag) =>
    formatTaxonomyFieldValue(tag.label, tag.termId),
  );
  const fieldValues = [
    wireValues.join(";"),
    wireValues.map((value) => `-1;#${value}`).join(";#"),
  ];

  let lastError = "Failed to update managed metadata";
  for (const fieldValue of fieldValues) {
    try {
      const result: any = await sp.web.lists
        .getByTitle(listNames.loan)
        .items.getById(itemId)
        .validateUpdateListItem(
          [{ FieldName: fieldInternalName, FieldValue: fieldValue }],
          true,
        );
      const failures = getValidateUpdateFailures(result);
      if (!failures.length) return;
      lastError =
        failures[0]?.ErrorMessage || failures[0]?.Message || lastError;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
  }
  throw new Error(lastError);
};

const findTermIdsByLabels = (
  nodes: TreeNode[],
  labels: string[],
): string[] => {
  if (!labels.length) return [];
  const normalizedLabels = new Set(
    labels.map((label) => label.trim().toLowerCase()),
  );
  const termIds: string[] = [];

  const walk = (treeNodes: TreeNode[]): void => {
    treeNodes.forEach((node) => {
      const label = String(node.label || node.data?.label || "")
        .trim()
        .toLowerCase();
      if (normalizedLabels.has(label)) {
        termIds.push(String(node.key));
      }
      if (node.children?.length) {
        walk(node.children);
      }
    });
  };

  walk(nodes);
  return termIds;
};

const buildTaxonomyDisplayValue = (
  termIds: string[],
  category: TaxonomyCategory,
  trees: Record<TaxonomyCategory, TreeNode[]>,
): string =>
  termIds
    .map((termId) => findNodeLabel(trees[category] || [], termId))
    .filter(Boolean)
    .join(", ");

const resolveTermIdsForCategory = (
  rawField: any,
  rowLabels: string,
  tree: TreeNode[],
): string[] => {
  const fromGuids = collectTaxonomyTermGuids(rawField);
  if (fromGuids.length) {
    return fromGuids;
  }
  return findTermIdsByLabels(tree, toTaxonomyLabels(rowLabels));
};

const updateTaxonomyFieldOnListItem = async (
  itemId: number,
  fieldInternalName: string,
  tags: ITaxonomyTag[],
): Promise<void> => {
  if (!fieldInternalName) return;

  if (!tags.length) {
    const result: any = await sp.web.lists
      .getByTitle(listNames.loan)
      .items.getById(itemId)
      .validateUpdateListItem(
        [{ FieldName: fieldInternalName, FieldValue: "" }],
        true,
      );
    const failures = getValidateUpdateFailures(result);
    if (failures.length) {
      throw new Error(
        failures[0]?.ErrorMessage ||
          failures[0]?.Message ||
          "Failed to clear tags",
      );
    }
    return;
  }

  await applyTaxonomyToListItem(itemId, tags);
};

const detectTaxonomyCategoryFromPath = (
  folderPath: string,
): TaxonomyCategory | null => {
  if (!folderPath) return null;
  if (folderPath.includes("Asset Management")) return "Asset Management";
  if (folderPath.includes("CSP Legal")) return "Legal";
  if (folderPath.includes("Servicing")) return "Servicing";
  return null;
};


/* eslint-disable @typescript-eslint/no-floating-promises */

// Multi-select managed metadata tag picker (dropdown + chevron).
const TaxonomyTagPicker = ({
  category,
  tree,
  value,
  displayLabels: displayLabelsProp,
  placeholder = "Select tags",
  disabled = false,
  onChange,
}: ITaxonomyTagPickerProps): React.ReactElement => {
  const rootRef = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const [overlayVisible, setOverlayVisible] = useState(false);
  const [selectedTermIds, setSelectedTermIds] = useState<string[]>(value);
  const [selectedLabels, setSelectedLabels] = useState<string[]>([]);
  const selectedTermIdsRef = useRef<string[]>(value);
  const selectedLabelsRef = useRef<string[]>([]);

  const syncSelectionFromValue = useCallback(
    (termIds: string[], labelsOverride?: string[]): void => {
      selectedTermIdsRef.current = termIds;
      setSelectedTermIds(termIds);
      if (!termIds.length) {
        selectedLabelsRef.current = [];
        setSelectedLabels([]);
        return;
      }

      const labels: string[] = [];
      termIds.forEach((termId, index) => {
        labels.push(
          labelsOverride?.[index] ||
            displayLabelsProp?.[index] ||
            findNodeLabel(tree, termId) ||
            "",
        );
      });
      const resolvedLabels = labels.filter(Boolean);
      selectedLabelsRef.current = resolvedLabels;
      setSelectedLabels(resolvedLabels);
    },
    [displayLabelsProp, tree],
  );

  useEffect(() => {
    syncSelectionFromValue(value);
  }, [value, syncSelectionFromValue]);

  useEffect(() => {
    if (disabled || !category) {
      setOverlayVisible(false);
    }
  }, [category, disabled]);

  useEffect(() => {
    if (!overlayVisible) return;
    const onDocumentMouseDown = (event: MouseEvent): void => {
      const target = event.target as HTMLElement | null;
      if (rootRef.current?.contains(target)) {
        return;
      }
      setOverlayVisible(false);
    };
    document.addEventListener("mousedown", onDocumentMouseDown);
    return () => document.removeEventListener("mousedown", onDocumentMouseDown);
  }, [overlayVisible]);

  const resolvedLabel =
    value.length === 0
      ? ""
      : value
          .map((termId) => findNodeLabel(tree, termId))
          .filter(Boolean)
          .join(", ") ||
        (displayLabelsProp || []).join(", ") ||
        selectedLabels.join(", ");
  const hasSelection = value.length > 0;

  const toggleOverlay = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    event.stopPropagation();
    if (disabled || !category) return;
    setOverlayVisible((prev) => !prev);
  };

  const handleToggle = (termId: string, label: string): void => {
    const currentTermIds = selectedTermIdsRef.current;
    const currentLabels = selectedLabelsRef.current;
    const isSelected = currentTermIds.some((id) => termIdsMatch(id, termId));
    const nextIds = isSelected
      ? currentTermIds.filter((id) => !termIdsMatch(id, termId))
      : [...currentTermIds, termId];

    syncSelectionFromValue(
      nextIds,
      isSelected
        ? currentLabels.filter(
            (_, index) => !termIdsMatch(currentTermIds[index], termId),
          )
        : [...currentLabels, label],
    );
    onChange(nextIds);
  };

  const nodeTemplate = (node: TreeNode): React.ReactElement => {
    const termId = String(node.key);
    const label = String(node.label || node.data?.label || "");
    const isSelected = selectedTermIds.some((id) => termIdsMatch(id, termId));

    return (
      <span
        role="button"
        tabIndex={0}
        data-selected={isSelected ? "true" : "false"}
        className={styles.copyMoveNode}
        onClick={(e) => {
          e.preventDefault();
          e.stopPropagation();
          handleToggle(termId, label);
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            e.stopPropagation();
            handleToggle(termId, label);
          }
        }}
      >
        <i className={`pi pi-tag ${styles.uploadTagIcon}`} />
        <span className={styles.copyMoveNodeLabel} title={label}>
          {label}
        </span>
        <i
          className={`pi pi-times ${styles.taxonomyTagRemoveIcon}`}
          title={isSelected ? "Remove tag" : undefined}
          aria-hidden={!isSelected}
          aria-label={isSelected ? `Remove ${label}` : undefined}
        />
      </span>
    );
  };

  return (
    <div
      ref={rootRef}
      className={styles.taxonomyTagPickerRoot}
      data-open={overlayVisible ? "true" : "false"}
    >
      <button
        ref={triggerRef}
        type="button"
        className={styles.taxonomyTagPicker}
        data-has-selection={hasSelection ? "true" : "false"}
        data-category-selected={category ? "true" : "false"}
        data-open={overlayVisible ? "true" : "false"}
        disabled={disabled || !category}
        aria-expanded={overlayVisible}
        onClick={toggleOverlay}
      >
        <span
          className={
            hasSelection ? styles.taxonomyTagValue : styles.uploadPathPlaceholder
          }
          title={hasSelection ? resolvedLabel : undefined}
        >
          {!category
            ? "Select a tag category first"
            : !tree.length
              ? `No ${category} terms found`
              : resolvedLabel || placeholder}
        </span>
        <i
          className={`pi ${
            overlayVisible ? "pi-angle-up" : "pi-angle-down"
          } ${styles.uploadPathPickerIcon}`}
        />
      </button>

      {overlayVisible && (
        <div className={styles.taxonomyTagDropdown}>
          <div className={styles.taxonomyOverlayTitle}>
            <span>{category ? `${category} tags` : "Tags"}</span>
            <span className={styles.taxonomyOverlayHint}>
              Click again to remove
            </span>
          </div>
          <div className={styles.taxonomyTreePanel}>
            {tree.length ? (
              <Tree
                className={styles.taxonomyTagTree}
                value={tree}
                nodeTemplate={nodeTemplate}
                propagateSelectionUp={false}
                propagateSelectionDown={false}
              />
            ) : (
              <div className={styles.copyMoveStatus}>No terms available</div>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

const Loan = (props: ILoanProps) => {
  const dispatch: any = useDispatch();
  const [loader, setLoader] = useState(false);
  const [drpdown, setDrpdown] = useState<IAllDropdowns>({ sponsor: [] });
  const [mainFolders, setMainFolders] = useState<ILoanRecord[]>([]);

  // Store filtered rows currently rendered in the table.
  const [displayItems, setDisplayItems] = useState<ILoanRecord[]>([]);

  // Track current folder path where empty means dashboard root.
  const [currentFolder, setCurrentFolder] = useState("");

  // Keep per-level search and sponsor filter state.
  const [filter, setFilter] = useState<ILoanFilterState>(EMPTY_FILTER);

  // Cache folder content by path to avoid repeated SharePoint fetches.
  // Populated on first visit to each folder. Lets Search/Filter/Reset and
  // breadcrumb back-navigation work instantly without re-hitting SharePoint.
  const folderCacheRef = useRef<Map<string, ILoanRecord[]>>(new Map());
  const loanTreeCacheRef = useRef<Map<string, ILoanRecord[]>>(new Map());
  const loanTreeLoadPromisesRef = useRef<Map<string, Promise<ILoanRecord[]>>>(
    new Map(),
  );
  const selectedLoanRootRef = useRef("");
  const [loanTreeSearchLoading, setLoanTreeSearchLoading] = useState(false);

  const clearFolderCaches = (): void => {
    folderCacheRef.current.clear();
    loanTreeCacheRef.current.clear();
    loanTreeLoadPromisesRef.current.clear();
  };

  const syncSelectedLoanRoot = (folderPath: string): void => {
    const loanRoot = getLoanRootPath(folderPath, mainFolders);
    selectedLoanRootRef.current = loanRoot
      ? normalizeSharePointPath(loanRoot)
      : "";
  };

  const getActiveLoanRoot = (folderPath: string): string | null => {
    if (selectedLoanRootRef.current) {
      return selectedLoanRootRef.current;
    }

    const loanRoot = getLoanRootPath(folderPath, mainFolders);
    return loanRoot ? normalizeSharePointPath(loanRoot) : null;
  };

  const getSelectedLoanLabel = (): string => {
    const loanRoot = getActiveLoanRoot(currentFolder);
    if (!loanRoot) {
      return "";
    }

    const loanRecord = mainFolders.find((loan) =>
      pathsEqual(loan.fileRef, loanRoot),
    );
    if (loanRecord?.fileName) {
      return loanRecord.fileName;
    }

    const parts = loanRoot.split("/").filter(Boolean);
    return parts[parts.length - 1] || "";
  };

  // Source file path highlighted after "Go to Parent File" navigation.
  const [highlightedOriginalFilePath, setHighlightedOriginalFilePath] =
    useState("");

  const menu = useRef<Menu>(null);
  const menuRef = useRef<Menu>(null);
  const selectedRowRef = useRef<any>(null);

  const [rowMenuItems, setRowMenuItems] = useState<any[]>([]);
  const [selectedFolder, setSelectedFolder] = useState<any>(null);
  const [showRename, setShowRename] = useState(false);
  const [folderName, setFolderName] = useState("");
  const [showLinkDialog, setShowLinkDialog] = useState(false);
  const [showVersionHistory, setShowVersionHistory] = React.useState(false);
  const [versionHistory, setVersionHistory] = React.useState<IVersionHistoryRow[]>(
    [],
  );
  const [versionHistoryTarget, setVersionHistoryTarget] =
    useState<ILoanRecord | null>(null);
  const [versionHistoryLoading, setVersionHistoryLoading] = useState(false);
  const [versionActionDialog, setVersionActionDialog] =
    useState<IVersionActionDialog>(EMPTY_VERSION_ACTION);
  const [versionActionLoading, setVersionActionLoading] = useState(false);
  const [versionDetails, setVersionDetails] =
    useState<IVersionHistoryRow | null>(null);
  const [showCreateFolder, setShowCreateFolder] = useState(false);
  const [isRequestFolderDialog, setIsRequestFolderDialog] = useState(false);
  const isRequestFolderDialogRef = useRef(false);
  const [requestFolderLoading, setRequestFolderLoading] = useState(false);
  const [showFileUpload, setShowFileUpload] = useState(false);
  const [uploadFiles, setUploadFiles] = useState<File[]>([]);
  const [uploadDragOver, setUploadDragOver] = useState(false);
  const [uploadTagCategory, setUploadTagCategory] =
    useState<TaxonomyCategory | null>(null);
  const [uploadTagTermIds, setUploadTagTermIds] = useState<string[]>([]);
  const [uploadTagTrees, setUploadTagTrees] = useState<
    Record<TaxonomyCategory, TreeNode[]>
  >({} as Record<TaxonomyCategory, TreeNode[]>);
  const [uploadTagFields, setUploadTagFields] = useState<
    Record<TaxonomyCategory, string>
  >({} as Record<TaxonomyCategory, string>);
  const [uploadTagsLoading, setUploadTagsLoading] = useState(false);
  const [showEditTagDialog, setShowEditTagDialog] = useState(false);
  const [editTagTarget, setEditTagTarget] = useState<ILoanRecord | null>(null);
  const [editTagCategory, setEditTagCategory] =
    useState<TaxonomyCategory | null>(null);
  const [editTagLoading, setEditTagLoading] = useState(false);
  const [editTagSaving, setEditTagSaving] = useState(false);
  const [editTagTermIds, setEditTagTermIds] = useState<
    Record<TaxonomyCategory, string[]>
  >({ ...EMPTY_EDIT_TAG_TERM_IDS });
  const [editTagTrees, setEditTagTrees] = useState<
    Record<TaxonomyCategory, TreeNode[]>
  >({} as Record<TaxonomyCategory, TreeNode[]>);
  const [editTagFields, setEditTagFields] = useState<
    Record<TaxonomyCategory, string>
  >({} as Record<TaxonomyCategory, string>);
  const [fileUploading, setFileUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState({
    completed: 0,
    total: 0,
    currentFileName: "",
    successCount: 0,
    failCount: 0,
    phase: "uploading" as "uploading" | "complete",
  });
  const fileUploadInputRef = useRef<HTMLInputElement>(null);
  const preserveFileUploadStateOnOpenRef = useRef(false);
  const [showDeleteDialog, setshowDeleteDialog] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<ILoanRecord | null>(null);
  const [deleteLoading, setDeleteLoading] = useState(false);
  const [showShareDialog, setShowShareDialog] = useState(false);
  const [shareDialogUrl, setShareDialogUrl] = useState("");
  const [shareItemName, setShareItemName] = useState("");
  const [folderTree, setFolderTree] = useState<TreeNode[]>([]);
  const [copyMoveExpandedKeys, setCopyMoveExpandedKeys] = useState<
    Record<string, boolean>
  >({});
  const [selectedDestination, setSelectedDestination] =
    useState<IFolderDestination | null>(null);
  const [copyMoveTreeLoading, setCopyMoveTreeLoading] = useState(false);
  const [showCopyMoveDialog, setShowCopyMoveDialog] = useState(false);
  const [showHyperlinkDialog, setShowHyperlinkDialog] = useState(false);
  const [hyperlinkTarget, setHyperlinkTarget] = useState<ILoanRecord | null>(
    null,
  );
  const [hyperlinkSourceLoan, setHyperlinkSourceLoan] =
    useState<ILoanRecord | null>(null);
  const [hyperlinkLoanOptions, setHyperlinkLoanOptions] = useState<ILoanRecord[]>(
    [],
  );
  const [selectedHyperlinkLoans, setSelectedHyperlinkLoans] = useState<
    ILoanRecord[]
  >([]);
  const [hyperlinkLoading, setHyperlinkLoading] = useState(false);
  const [copyMoveOperation, setCopyMoveOperation] = useState<"copy" | "move">(
    "copy",
  );
  const [copyMoveTarget, setCopyMoveTarget] = useState<ILoanRecord | null>(
    null,
  );
  const copyMoveTargetRef = useRef<ILoanRecord | null>(null);
  const shareIframeRef = useRef<HTMLIFrameElement>(null);
  const copyMoveOperationRef = useRef<"copy" | "move">("copy");
  const copyMoveLibraryRootRef = useRef<IFolderDestination | null>(null);
  const isAdmin: boolean = useSelector((e: any) => e.MainSPContext.isAdmin);

  // Folder depth after the document library root (0 = dashboard).
  // Example: library/1234/12345/123456 → depth 3
  const getFolderDepth = (): number => {
    if (!currentFolder) return 0;
    const parts = currentFolder.split("/").filter(Boolean);
    const libraryIndex = parts.indexOf(LIBRARY_NAME);
    if (libraryIndex < 0) return 0;
    return Math.max(0, parts.length - libraryIndex - 1);
  };

  // User-facing levels: CSP Loan Files = level 1, Loan Number = level 2.
  // Depth 0/1 → Request button; depth 2+ → existing New menu.
  const currentFolderDepth = getFolderDepth();
  const canShowNewButton = currentFolderDepth >= 2;
  const canShowRequestFolder =
    currentFolderDepth === 0 || currentFolderDepth === 1;

  const openRequestFolderDialog = (): void => {
    setFolderName("");
    isRequestFolderDialogRef.current = true;
    setIsRequestFolderDialog(true);
    setShowCreateFolder(true);
  };

  const newMenuItems = useMemo(
    () => [
      {
        label: "New sub folder",
        icon: "pi pi-folder-plus",
        visible: isAdmin,
        command: () => {
          setFolderName("");
          isRequestFolderDialogRef.current = false;
          setIsRequestFolderDialog(false);
          setShowCreateFolder(true);
        },
      },
      {
        label: "File Upload",
        icon: "pi pi-file",
        command: () => setShowFileUpload(true),
      },
    ],
    [isAdmin],
  );

  // Navigate a hyperlink row to the original source file's parent folder.
  const goToParentFile = (rowData: ILoanRecord): void => {
    const sourcePath = normalizeSharePointPath(
      rowData?.fileRef ||
        getServerRelativePathFromUrl(rowData?.serverRelativeUrl || ""),
    );
    if (!sourcePath) return;
    const parentFolder = getParentFolderPath(sourcePath);
    if (!parentFolder) return;
    void navigateToFolder(parentFolder, sourcePath);
  };

  const isHighlightedOriginalFile = (row: ILoanRecord): boolean =>
    !!highlightedOriginalFilePath &&
    pathsEqual(row.fileRef || row.serverRelativeUrl || "", highlightedOriginalFilePath);

  // Build row action menu items for the selected file or folder.
  const getRowMenuItems = (rowData: any): void => {
    if (rowData?.isHyperlink) {
      setRowMenuItems([
        {
          label: "Go to Parent File",
          icon: "pi pi-folder-open",
          command: () => goToParentFile(rowData),
        },
      ]);
      return;
    }

    setRowMenuItems([
      {
        label: "Open in App",
        icon: "pi pi-desktop",
        visible: selectedRowRef.current?.folderType !== 1,
        command: () => openFileInDesktopApp(selectedRowRef.current),
      },
      {
        label: "Download",
        icon: "pi pi-download",
        command: () => downloadFile(selectedRowRef.current),
      },
      {
        label: "Share",
        icon: "pi pi-share-alt",
        command: () => {
          openNativeShareDialog(selectedRowRef.current);
        },
      },
      {
        label: "Copy Link",
        icon: "pi pi-link",
        command: () => copySharePointLink(selectedRowRef.current),
      },
      {
        label: "Move To",
        visible: canShowNewButton,
        icon: "pi pi-arrow-right-arrow-left",
        command: () => {
          openCopyMoveDialog(selectedRowRef.current, "move");
        },
      },
      {
        label: "Copy To",
        icon: "pi pi-copy",
         visible: canShowNewButton,
        command: () => {
          openCopyMoveDialog(selectedRowRef.current, "copy");
        },
      },
      {
        label: "Hyperlink",
        icon: "pi pi-link",
        visible: selectedRowRef.current?.folderType !== 1,
        command: () => {
          openHyperlinkDialog(selectedRowRef.current);
        },
      },
      {
        label: "Rename",
        icon: "pi pi-pencil",
        visible: canShowNewButton,
        command: () => {
          const row = selectedRowRef.current;
          setSelectedFolder(row);
          setShowRename(true);
          // For files: strip extension so input only shows the base name
          // For folders: use name as-is
          const name = row.fileName || "";
          const isFile = row.folderType !== 1;
          setFolderName(
            isFile ? name.substring(0, name.lastIndexOf(".")) || name : name,
          );
        },
      },
      {
        label: "Edit Tag",
        icon: "pi pi-tags",
        visible:
          selectedRowRef.current?.folderType !== 1 &&
          !!detectTaxonomyCategoryFromPath(currentFolder),
        command: () => {
          void openEditTagDialog(selectedRowRef.current);
        },
      },
      {
        label: "Version History",
        icon: "pi pi-history",
        visible: selectedRowRef.current?.folderType !== 1,
        command: () => {
          getVersionHistory(selectedRowRef.current);
        },
      },
      {
        label: "Delete",
        icon: "pi pi-trash",
        className: "deleteMenu",
        visible: canShowNewButton,
        command: () => {
          setDeleteTarget(selectedRowRef.current);
          setshowDeleteDialog(true);
        },
      },
    ]);
  };
  // Load root-level loan folders and update dashboard state.
  const getLoanData = async (
    sponsorOptions: any[],
    showLoader = true,
  ): Promise<void> => {
    try {
      if (showLoader) {
        setLoader(true);
      }

      const list = sp.web.lists.getByTitle(listNames.loan);
      const rootFolder = await list.rootFolder();
      const rootFolderUrl = rootFolder.ServerRelativeUrl;

      const response: any = await list.renderListDataAsStream({
        ViewXml: `
          <View Scope="DefaultValue">
            <Query>
              <Where>
                <Eq>
                  <FieldRef Name="FSObjType" />
                  <Value Type="Integer">1</Value>
                </Eq>
              </Where>
              <OrderBy>
                <FieldRef Name="ID" Ascending="FALSE" />
              </OrderBy>
            </Query>
            <ViewFields>${VIEW_FIELDS_XML}</ViewFields>
            <RowLimit>2000</RowLimit>
          </View>
        `,
        FolderServerRelativeUrl: rootFolderUrl, // scopes query — avoids full recursive scan
      });

      const rootFolders: ILoanRecord[] = (response.Row || []).map(
        mapRowToLoanRecord,
      );

      setMainFolders(rootFolders);
      dispatch(setLoanDetails(rootFolders));
      setDisplayItems(applyFilter(rootFolders, filter));
      setDrpdown((prev) => ({ ...prev, sponsor: sponsorOptions }));

      // Root list is always fresh — clear any stale cached subfolder data
      clearFolderCaches();
    } catch (error) {
      console.error("getLoanData error:", error);
    } finally {
      if (showLoader) {
        setLoader(false);
      }
    }
  };
  // Load sponsor dropdown values and then load loan folders.
  const getSponsorData = async (): Promise<void> => {
    try {
      let paged = await sp.web.lists
        .getByTitle(listNames.sponsors)
        .items.select("*", "Author/Title")
        .expand("Author")
        .filter("IsDelete ne 1")
        .top(5000)
        .getPaged();

      const allRows = [...paged.results];
      while (paged.hasNext) {
        paged = await paged.getNext();
        allRows.push(...paged.results);
      }

      const sponsorOptions = allRows.map((item: any) => ({
        code: item.Id,
        name: item.Title,
      }));

      await getLoanData(sponsorOptions);
    } catch (error) {
      console.error("getSponsorData error:", error);
    }
  };

  // Load active virtual files mapped to the current destination loan folder.
  const fetchMappedFiles = async (
    folderServerRelativeUrl: string,
    options?: { allPathsForLoan?: boolean },
  ): Promise<ILoanRecord[]> => {
    const folderPath = normalizeSharePointPath(folderServerRelativeUrl);
    const destinationLoan =
      mainFolders
        .filter(
          (loan) =>
            !!loan.fileRef &&
            isFolderDescendantOrSelf(loan.fileRef, folderPath),
        )
        .sort(
          (left, right) =>
            normalizeSharePointPath(right.fileRef).length -
            normalizeSharePointPath(left.fileRef).length,
        )[0] || null;

    if (!destinationLoan?.id || !destinationLoan.fileRef) {
      return [];
    }

    const loanPath = normalizeSharePointPath(destinationLoan.fileRef);
    const relativeFolderPath = folderPath
      .substring(loanPath.length)
      .replace(/^\/+|\/+$/g, "");
    const mappingFilter = options?.allPathsForLoan
      ? `DestinationLoanItemId eq ${destinationLoan.id} and IsActive eq 1`
      : `DestinationLoanItemId eq ${destinationLoan.id} and ` +
        `RelativeFolderPath eq '${escapeODataString(relativeFolderPath)}' and ` +
        "IsActive eq 1";
    const mappings: ILoanFileMapping[] = await sp.web.lists
      .getByTitle(listNames.loanFileMapping)
      .items.select(
        "Id",
        "Title",
        "SourceFileItemId",
        "SourceFileUrl",
        "DestinationLoanItemId",
        "RelativeFolderPath",
        "Created",
        "Modified",
        "Author/Title",
      )
      .expand("Author")
      .filter(mappingFilter)
      .top(5000)()
      .catch((error: any) => {
        console.error("Load file hyperlinks failed:", error);
        return [];
      });

    // Pull Asset Mgmt / Legal / Servicing from the original source files.
    const sourceIds = Array.from(
      new Set(
        mappings
          .map((mapping) => Number(mapping.SourceFileItemId || 0))
          .filter((id) => id > 0),
      ),
    );
    const sourceMetadataById = new Map<
      number,
      Pick<ILoanRecord, "assetmanagement" | "servicing" | "legal" | "fileName">
    >();

    if (sourceIds.length) {
      const list = sp.web.lists.getByTitle(listNames.loan);
      const chunkSize = 30;

      for (let index = 0; index < sourceIds.length; index += chunkSize) {
        const chunk = sourceIds.slice(index, index + chunkSize);
        const valuesXml = chunk
          .map((id) => `<Value Type="Counter">${id}</Value>`)
          .join("");

        try {
          const response: any = await list.renderListDataAsStream({
            ViewXml: `
              <View Scope="RecursiveAll">
                <Query>
                  <Where>
                    <In>
                      <FieldRef Name="ID" />
                      <Values>${valuesXml}</Values>
                    </In>
                  </Where>
                </Query>
                <ViewFields>${VIEW_FIELDS_XML}</ViewFields>
                <RowLimit>${chunk.length}</RowLimit>
              </View>
            `,
          });

          (response.Row || []).forEach((row: any) => {
            const source = mapRowToLoanRecord(row);
            if (!source.id) return;
            sourceMetadataById.set(source.id, {
              assetmanagement: source.assetmanagement,
              servicing: source.servicing,
              legal: source.legal,
              fileName: source.fileName,
            });
          });
        } catch (error) {
          console.error("Load hyperlink source metadata failed:", error);
        }
      }
    }

    const seenSources = new Set<string>();
    return mappings.reduce<ILoanRecord[]>((records, mapping) => {
      const sourceUrl = getSharePointUrlValue(mapping.SourceFileUrl);
      const sourceId = Number(mapping.SourceFileItemId || 0);
      const sourceKey = `${sourceId}|${sourceUrl.toLowerCase()}`;

      if (!sourceUrl || seenSources.has(sourceKey)) {
        return records;
      }
      seenSources.add(sourceKey);

      const sourceMeta = sourceMetadataById.get(sourceId);
      const fallbackName = String(
        mapping.Title || sourceUrl.split("/").pop() || "Linked file",
      );

      records.push({
        id: sourceId || null,
        mappingId: Number(mapping.Id || 0) || undefined,
        name: sourceMeta?.fileName || fallbackName,
        fileName: sourceMeta?.fileName || fallbackName,
        fileRef: getServerRelativePathFromUrl(sourceUrl),
        serverRelativeUrl: sourceUrl,
        folderType: 0,
        sponsor: destinationLoan.sponsor,
        createdby: String(mapping.Author?.Title || ""),
        createddate: toSerializableDate(mapping.Created),
        modifieddate: toSerializableDate(mapping.Modified),
        assetmanagement: sourceMeta?.assetmanagement || "",
        servicing: sourceMeta?.servicing || "",
        legal: sourceMeta?.legal || "",
        type: "hyperlink",
        isHyperlink: true,
      });
      return records;
    }, []);
  };

  // Fetch direct child items for one folder path.
  const fetchFolderContents = async (
    folderServerRelativeUrl: string,
  ): Promise<ILoanRecord[]> => {
    const list = sp.web.lists.getByTitle(listNames.loan);
    const [response, mappedFiles]: [any, ILoanRecord[]] = await Promise.all([
      list.renderListDataAsStream({
        ViewXml: `
          <View Scope="DefaultValue">
            <Query>
              <OrderBy>
                <FieldRef Name="FSObjType" Ascending="FALSE" />
                <FieldRef Name="FileLeafRef" Ascending="TRUE" />
              </OrderBy>
            </Query>
            <ViewFields>${VIEW_FIELDS_XML}</ViewFields>
            <RowLimit>2000</RowLimit>
          </View>
        `,
        FolderServerRelativeUrl: folderServerRelativeUrl,
      }),
      fetchMappedFiles(folderServerRelativeUrl),
    ]);

    return [
      ...(response.Row || []).map(mapRowToLoanRecord),
      ...mappedFiles,
    ];
  };
  // Get root or folder base list using cache-first behavior.
  const getBaseListForLevel = async (
    folderPath: string,
  ): Promise<ILoanRecord[]> => {
    if (!folderPath) return mainFolders;

    if (folderCacheRef.current.has(folderPath)) {
      return folderCacheRef.current.get(folderPath)!;
    }
    const items = await fetchFolderContents(folderPath);
    folderCacheRef.current.set(folderPath, items);
    return items;
  };
  // Apply search and sponsor filters against a base list.
  const applyFilter = (
    base: ILoanRecord[],
    state: ILoanFilterState,
  ): ILoanRecord[] => {
    let result = [...base];

    if (state.search.trim()) {
      const keyword = state.search.toLowerCase();
      result = result.filter(
        (item) =>
          item.fileName?.toLowerCase().includes(keyword) ||
          item.createdby?.toLowerCase().includes(keyword) ||
          item.sponsor?.sponsorTitle?.toLowerCase().includes(keyword),
      );
    }

    if (state.sponsor) {
      result = result.filter(
        (item) => item.sponsor?.sponsorTitle === state.sponsor!.name,
      );
    }

    return result;
  };

  const invalidateLoanTreeCacheForPath = (folderPath: string): void => {
    const loanRoot = getLoanRootPath(folderPath, mainFolders);
    if (!loanRoot) {
      return;
    }

    const normalizedRoot = normalizeSharePointPath(loanRoot);
    loanTreeCacheRef.current.delete(normalizedRoot);
    loanTreeLoadPromisesRef.current.delete(normalizedRoot);
  };

  const fetchLoanFolderTreeItems = async (
    loanRootPath: string,
  ): Promise<ILoanRecord[]> => {
    const normalizedRoot = normalizeSharePointPath(loanRootPath);
    const list = sp.web.lists.getByTitle(listNames.loan);
    const rows: ILoanRecord[] = [];
    let pagingToken: string | undefined;

    do {
      const response: any = await list.renderListDataAsStream({
        ViewXml: `
          <View Scope="RecursiveAll">
            <Query>
              <OrderBy>
                <FieldRef Name="FSObjType" Ascending="FALSE" />
                <FieldRef Name="FileLeafRef" Ascending="TRUE" />
              </OrderBy>
            </Query>
            <ViewFields>${VIEW_FIELDS_XML}</ViewFields>
            <RowLimit Paged="TRUE">2000</RowLimit>
          </View>
        `,
        FolderServerRelativeUrl: normalizedRoot,
        Paging: pagingToken,
      });

      rows.push(...(response.Row || []).map(mapRowToLoanRecord));
      pagingToken = response.NextHref
        ? response.NextHref.split("?")[1]
        : undefined;
    } while (pagingToken);

    const mappedFiles = await fetchMappedFiles(normalizedRoot, {
      allPathsForLoan: true,
    });

    return dedupeLoanRecords(
      scopeItemsToLoanRoot(
        [...rows, ...mappedFiles],
        normalizedRoot,
        mainFolders,
      ),
    );
  };

  const ensureLoanTreeLoaded = async (
    loanRootPath: string,
  ): Promise<ILoanRecord[]> => {
    const normalizedRoot = normalizeSharePointPath(loanRootPath);
    const cached = loanTreeCacheRef.current.get(normalizedRoot);
    if (cached) {
      return cached;
    }

    const inFlight = loanTreeLoadPromisesRef.current.get(normalizedRoot);
    if (inFlight) {
      return inFlight;
    }

    const promise = fetchLoanFolderTreeItems(normalizedRoot)
      .then((items) => {
        loanTreeCacheRef.current.set(normalizedRoot, items);
        loanTreeLoadPromisesRef.current.delete(normalizedRoot);
        return items;
      })
      .catch((error) => {
        loanTreeLoadPromisesRef.current.delete(normalizedRoot);
        throw error;
      });

    loanTreeLoadPromisesRef.current.set(normalizedRoot, promise);
    return promise;
  };

  const resolveDisplayItems = async (
    folderPath: string,
    state: ILoanFilterState,
  ): Promise<ILoanRecord[]> => {
    if (!folderPath) {
      return applyFilter(mainFolders, state);
    }

    const loanRoot = getActiveLoanRoot(folderPath);
    if (loanRoot && state.search.trim()) {
      const cachedTree = loanTreeCacheRef.current.get(loanRoot);
      const treeItems = cachedTree || (await ensureLoanTreeLoaded(loanRoot));
      return filterLoanTreeItems(treeItems, state, loanRoot, mainFolders);
    }

    const base = await getBaseListForLevel(folderPath);
    return applyFilter(base, state);
  };

  const refreshCurrentView = async (): Promise<void> => {
    setDisplayItems(await resolveDisplayItems(currentFolder, filter));
  };
  // Navigate to any folder level and refresh visible rows.
  const navigateToFolder = async (
    folderPath: string,
    highlightOriginalPath = "",
  ): Promise<void> => {
    try {
      setLoader(true);
      setHighlightedOriginalFilePath(
        highlightOriginalPath
          ? normalizeSharePointPath(highlightOriginalPath)
          : "",
      );
      syncSelectedLoanRoot(folderPath);
      setCurrentFolder(folderPath);
      setDisplayItems(await resolveDisplayItems(folderPath, filter));

      const loanRoot = getActiveLoanRoot(folderPath);
      if (loanRoot) {
        void ensureLoanTreeLoaded(loanRoot).catch((error) => {
          console.error("Preload loan tree failed:", error);
        });
      }
    } catch (error) {
      console.error("navigateToFolder error:", error);
    } finally {
      setLoader(false);
    }
  };
  // Navigate into folders or open files in a new browser tab (no download).
  const handleFolderClick = (row: ILoanRecord): void => {
    if (row.folderType === 1) {
      navigateToFolder(row.fileRef);
      return;
    }

    const absoluteUrl = toAbsoluteSharePointUrl(
      row.serverRelativeUrl || row.fileRef,
    );
    if (!absoluteUrl) {
      toastFunc("warn", "Warning", "File URL is unavailable");
      return;
    }

    const browserUrl = toBrowserPreviewUrl(absoluteUrl, row.fileName || "");
    window.open(browserUrl, "_blank", "noopener,noreferrer");
  };
  // Navigate breadcrumb back to dashboard root.
  const goToRoot = (): void => {
    navigateToFolder("");
  };
  // Navigate breadcrumb to any intermediate folder segment.
  const goToFolder = (folderPath: string): void => {
    navigateToFolder(folderPath);
  };
  // Update search filter — instant at root; in-memory tree filter inside a loan.
  const onSearchInputChange = (value: string): void => {
    const newFilter = { ...filter, search: value };
    setFilter(newFilter);

    if (!currentFolder) {
      setDisplayItems(applyFilter(mainFolders, newFilter));
      return;
    }

    const loanRoot = getActiveLoanRoot(currentFolder);
    const trimmedSearch = value.trim();

    if (loanRoot && trimmedSearch) {
      const cachedTree = loanTreeCacheRef.current.get(loanRoot);
      if (cachedTree) {
        setDisplayItems(
          filterLoanTreeItems(cachedTree, newFilter, loanRoot, mainFolders),
        );
        return;
      }

      void (async () => {
        setLoanTreeSearchLoading(true);
        try {
          const treeItems = await ensureLoanTreeLoaded(loanRoot);
          setDisplayItems(
            filterLoanTreeItems(treeItems, newFilter, loanRoot, mainFolders),
          );
        } catch (error) {
          console.error("Loan tree search failed:", error);
        } finally {
          setLoanTreeSearchLoading(false);
        }
      })();
      return;
    }

    const cachedLevel = folderCacheRef.current.get(currentFolder);
    if (cachedLevel) {
      setDisplayItems(applyFilter(cachedLevel, newFilter));
      return;
    }

    void (async () => {
      const base = await getBaseListForLevel(currentFolder);
      setDisplayItems(applyFilter(base, newFilter));
    })();
  };

  // Update sponsor filter and refresh current level rows.
  const onSponsorChange = async (
    value: IDrpdownOptions | null,
  ): Promise<void> => {
    const newFilter = { ...filter, sponsor: value };
    setFilter(newFilter);
    setDisplayItems(await resolveDisplayItems(currentFolder, newFilter));
  };

  // Reset filters and restore unfiltered current level rows.
  const onReset = async (): Promise<void> => {
    setFilter(EMPTY_FILTER);
    setDisplayItems(await resolveDisplayItems(currentFolder, EMPTY_FILTER));
  };
  // Build breadcrumb segments from current folder path.
  const getBreadcrumbSegments = (): { label: string; path: string }[] => {
    if (!currentFolder) return [];
    const parts = currentFolder.split("/").filter(Boolean);
    const libraryIndex = parts.indexOf(LIBRARY_NAME);
    const folders = libraryIndex >= 0 ? parts.slice(libraryIndex + 1) : [];
    return folders.map((folder, index) => ({
      label: folder,
      path: "/" + parts.slice(0, libraryIndex + index + 2).join("/"),
    }));
  };

  // Build avatar initials from a display name.
  const getInitials = (name: string): string => {
    if (!name) return "?";
    const parts = name.trim().split(" ");
    return parts.length >= 2
      ? `${parts[0][0]}${parts[parts.length - 1][0]}`.toUpperCase()
      : name.slice(0, 2).toUpperCase();
  };
  // Render Name column with icon and click behavior.
  const nameBodyTemplate = (row: ILoanRecord): React.ReactElement => {
    const isOriginal = isHighlightedOriginalFile(row);

    return (
      <span
        className={`${styles.nameCell} ${
          isOriginal ? styles.originalFileNameCell : ""
        }`}
        onClick={() => handleFolderClick(row)}
        title={
          isOriginal
            ? `Original file: ${row.fileName}`
            : row.fileName
        }
      >
        {row.folderType === 1 ? (
          <i className={`pi pi-folder ${styles.folderIcon}`} />
        ) : (
          <i className={`${getFileIcon(row.fileName)} ${styles.fileIcon}`} />
        )}
        <span className={styles.truncateCell}>{row.fileName}</span>
        {isOriginal && (
          <span className={styles.originalFileBadge} title="Original source file">
            <i className="pi pi-bookmark-fill" />
            Original
          </span>
        )}
      </span>
    );
  };
  // Render Created By column with avatar initials and name.
  const createdByBodyTemplate = (rowData: ILoanRecord): React.ReactElement => (
    <div className={styles.createdByCell} title={rowData.createdby || ""}>
      <span className={styles.avatar} style={{ backgroundColor: AVATAR_COLOR }}>
        {getInitials(rowData.createdby ?? "")}
      </span>
      <span className={styles.createdByName}>{rowData.createdby}</span>
    </div>
  );
  // Format date values for table columns.
  const createdOnBodyTemplate = (rowData: ILoanRecord): React.ReactElement => (
    <span>{formatDateTime(rowData.createddate)}</span>
  );
  const modifiedOnBodyTemplate = (rowData: ILoanRecord): React.ReactElement => (
    <span>{formatDateTime(rowData.modifieddate)}</span>
  );

  const metadataPillCell = (value?: string): React.ReactElement => {
    const labels = toTaxonomyLabels(value);

    if (!labels.length) {
      return <span className={styles.metadataTagEmpty}>—</span>;
    }

    return (
      <div
        className={`${styles.metadataTagPills} ${styles.metadataTableTagPills}`}
        title={labels.join(", ")}
      >
        {labels.map((label, index) => (
          <span
            key={`${label}-${index}`}
            className={`${styles.metadataTagPill} ${styles.metadataTableTagPill}`}
            title={label}
          >
            <i className={`pi pi-tag ${styles.metadataTableTagIcon}`} />
            <span className={styles.metadataTableTagText}>{label}</span>
          </span>
        ))}
      </div>
    );
  };

  const shouldSkipDownloadFolder = (name: string): boolean =>
    name.startsWith("_") || name === "Forms";

  // Download a file directly or a folder as recursive zip.
  const downloadFile = async (rowData: any): Promise<void> => {
    if (!rowData?.fileRef) {
      toastFunc("error", "Error", "Invalid item selected");
      return;
    }

    try {
      setLoader(true);
      const isFile = rowData.folderType !== 1;

      if (isFile) {
        const buffer = await sp.web
          .getFileByServerRelativePath(rowData.fileRef)
          .getBuffer();
        triggerBlobDownload(new Blob([buffer]), rowData.fileName);
        return;
      }

      // eslint-disable-next-line @typescript-eslint/no-var-requires
      const JSZipModule = require("jszip");
      const zip = new JSZipModule();

      const addFolderToZip = async (
        folderPath: string,
        zipFolder: typeof zip,
      ): Promise<void> => {
        const [files, subFolders] = await Promise.all([
          sp.web.getFolderByServerRelativePath(folderPath).files(),
          sp.web.getFolderByServerRelativePath(folderPath).folders(),
        ]);

        await Promise.all(
          files.map(async (file) => {
            const buffer = await sp.web
              .getFileByServerRelativePath(file.ServerRelativeUrl)
              .getBuffer();
            zipFolder.file(file.Name, buffer);
          }),
        );

        await Promise.all(
          subFolders
            .filter((subFolder) => !shouldSkipDownloadFolder(subFolder.Name))
            .map(async (subFolder) => {
              const childZipFolder = zipFolder.folder(subFolder.Name);
              if (childZipFolder) {
                await addFolderToZip(subFolder.ServerRelativeUrl, childZipFolder);
              }
            }),
        );
      };

      await addFolderToZip(rowData.fileRef, zip);
      const zipBlob: Blob = await zip.generateAsync({ type: "blob" });
      triggerBlobDownload(zipBlob, `${rowData.fileName}.zip`);
    } catch (err) {
      console.error("downloadFile error:", err);
      toastFunc("error", "Error", "Failed to download");
    } finally {
      setLoader(false);
    }
  };
  // Copy SharePoint link for the selected item to clipboard.
  const copySharePointLink = async (rowData: any): Promise<void> => {
    setLoader(true);
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const itemId = rowData.id;
    const listName = listNames.loan;

    const requestDigest = await fetch(`${siteUrl}/_api/contextinfo`, {
      method: "POST",
      headers: { Accept: "application/json;odata=verbose" },
    })
      .then((res) => res.json())
      .then((data) => data.d.GetContextWebInformation.FormDigestValue);

    const linkKinds = [
      { linkKind: 1, role: 1 },
      { linkKind: 2, role: 2 },
      { linkKind: 3, role: 1 },
      { linkKind: 4, role: 2 },
    ];

    for (const link of linkKinds) {
      const response = await fetch(
        `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})/ShareLink`,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": requestDigest,
          },
          body: JSON.stringify({
            request: {
              createLink: true,
              settings: {
                linkKind: link.linkKind,
                role: link.role,
                expiration: null,
              },
            },
          }),
        },
      ).then((res) => res.json());

      const sharingLink = response?.d?.ShareLink?.sharingLinkInfo?.Url;
      if (sharingLink) {
        try {
          await navigator.clipboard.writeText(sharingLink);
        } catch {
          const textArea = document.createElement("textarea");
          textArea.value = sharingLink;
          document.body.appendChild(textArea);
          textArea.select();
          document.execCommand("copy");
          document.body.removeChild(textArea);
        }

        setLoader(false);
        setShowLinkDialog(true);
        return;
      }
    }

    setLoader(false);
    toastFunc("error", "Error", "Failed to copy link");
  };
  //Rename folder
  const resolveLoanNumberFromPath = (itemPath: string): string => {
    const path = normalizeSharePointPath(itemPath);
    if (!path) return "";

    const loan =
      mainFolders
        .filter(
          (folder) =>
            !!folder.fileRef &&
            isFolderDescendantOrSelf(folder.fileRef, path),
        )
        .sort(
          (left, right) =>
            normalizeSharePointPath(right.fileRef).length -
            normalizeSharePointPath(left.fileRef).length,
        )[0] || null;

    return loan?.fileName || "";
  };

  const logActivity = async (params: {
    title: string;
    action: string;
    fromPath: string;
    toPath?: string;
    loanNumber?: string;
    result: "Success" | "Failed";
  }): Promise<void> => {
    try {
      const fromPath = normalizeSharePointPath(params.fromPath).substring(
        0,
        255,
      );
      const toPath = normalizeSharePointPath(params.toPath || "").substring(
        0,
        255,
      );
      const loanNumber = (
        params.loanNumber ||
        resolveLoanNumberFromPath(fromPath || toPath)
      ).substring(0, 255);

      await sp.web.lists.getByTitle(listNames.loanActivityLog).items.add({
        Title: String(params.title || "").substring(0, 255),
        Action: String(params.action || "").substring(0, 255),
        FromPath: fromPath,
        ToPath: toPath,
        LoanNumber: loanNumber,
        Result: params.result,
      });
    } catch (error) {
      console.warn("LoanActivityLog write failed:", error);
    }
  };

  const renameFolder = async (
    folderPath: string,
    newName: string,
  ): Promise<void> => {
    if (!newName.trim()) {
      toastFunc("warn", "Warning", "Please enter a name");
      return;
    }

    const isFile = selectedFolder?.folderType !== 1;
    const extension = isFile
      ? `.${selectedFolder?.fileName?.split(".").pop()}`
      : "";
    const fullNewName = `${newName.trim()}${extension}`;
    const oldName = folderPath.substring(folderPath.lastIndexOf("/") + 1);

    if (fullNewName === oldName) {
      setShowRename(false);
      return;
    }

    try {
      setLoader(true);
      const parentPath = folderPath.substring(0, folderPath.lastIndexOf("/"));
      const newFolderPath = `${parentPath}/${fullNewName}`;

      if (isFile) {
        // Resolve list item ID before updating file name.
        const fileItem: any = await sp.web
          .getFileByServerRelativePath(folderPath)
          .listItemAllFields();

        // Rename file using validateUpdateListItem for version-safe update.
        await sp.web.lists
          .getByTitle(listNames.loan)
          .items.getById(fileItem.Id)
          .validateUpdateListItem(
            [{ FieldName: "FileLeafRef", FieldValue: fullNewName }],
            false,
          );
      } else {
        // Rename folder by moveByPath and then update folder title.
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .moveByPath(newFolderPath, true);

        const folderItem: any = await sp.web
          .getFolderByServerRelativePath(newFolderPath)
          .listItemAllFields();

        await sp.web.lists
          .getByTitle(listNames.loan)
          .items.getById(folderItem.Id)
          .update({ Title: fullNewName });
      }

      // Apply optimistic UI updates after rename operation.
      const levelPath = currentFolder || parentPath;
      const cachedItems = folderCacheRef.current.get(levelPath) ?? mainFolders;

      const updatedItems = cachedItems.map((item) =>
        item.fileRef === folderPath
          ? {
              ...item,
              fileName: fullNewName,
              name: fullNewName,
              fileRef: newFolderPath,
              serverRelativeUrl: newFolderPath,
            }
          : item,
      );

      folderCacheRef.current.delete(folderPath);
      folderCacheRef.current.set(levelPath, updatedItems);
      invalidateLoanTreeCacheForPath(folderPath);

      if (filter.search.trim() && getActiveLoanRoot(currentFolder)) {
        void refreshCurrentView();
      } else {
        setDisplayItems(applyFilter(updatedItems, filter));
      }

      if (!currentFolder) {
        setMainFolders(updatedItems);
        dispatch(setLoanDetails(updatedItems));
      }

      await logActivity({
        title: fullNewName,
        action: "Rename",
        fromPath: folderPath,
        toPath: newFolderPath,
        result: "Success",
      });

      toastFunc(
        "success",
        "Success",
        `Renamed to "${fullNewName}" successfully`,
      );
      setShowRename(false);
      setSelectedFolder(null);
      setFolderName("");
    } catch (error) {
      console.error("renameFolder error:", error);
      await logActivity({
        title: selectedFolder?.fileName || newName.trim(),
        action: "Rename",
        fromPath: folderPath,
        toPath: "",
        result: "Failed",
      });
      toastFunc("error", "Error", "Failed to rename");
    } finally {
      setLoader(false);
    }
  };
  //open in app
  const openFileInDesktopApp = (row: any) => {
    const fileUrl = toAbsoluteSharePointUrl(row.serverRelativeUrl);
    const extension = row.fileName.split(".").pop()?.toLowerCase();
    switch (extension) {
      case "doc":
      case "docx":
        window.location.href = `ms-word:ofe|u|${fileUrl}`;
        break;

      case "xls":
      case "xlsx":
      case "xlsm":
        window.location.href = `ms-excel:ofe|u|${fileUrl}`;
        break;

      case "ppt":
      case "pptx":
        window.location.href = `ms-powerpoint:ofe|u|${fileUrl}`;
        break;

      default:
        window.open(fileUrl, "_blank");
        break;
    }
  };
  // Load version history for the selected file.
  const getVersionHistory = async (
    rowData: ILoanRecord,
    showDialogLoader = true,
  ): Promise<void> => {
    try {
      setVersionHistoryTarget(rowData);
      setShowVersionHistory(true);
      if (showDialogLoader) {
        setVersionHistory([]);
        setVersionHistoryLoading(true);
      }

      const file = sp.web.getFileByServerRelativePath(rowData.fileRef);
      const item = await file.getItem();

      const [
        fileInfo,
        fileVersions,
        itemVersions,
        currentFieldValuesAsText,
        webInfo,
      ] = await Promise.all([
          item
            .select(
              "*",
              "Author/Title",
              "Editor/Title",
              `${loanLibraryFields.sponsorName}/Title`,
              `${loanLibraryFields.sponsorName}/Id`,
            )
            .expand("Author", "Editor", loanLibraryFields.sponsorName)(),
          file.versions(),
          item.versions(),
          item.fieldValuesAsText(),
          sp.web.select("ServerRelativeUrl")(),
        ]);

      const itemId = Number(fileInfo.Id ?? fileInfo.ID ?? rowData.id ?? 0);
      const webServerRelativeUrl = String(
        webInfo?.ServerRelativeUrl || "",
      );
      const fileVersionByLabel = new Map<string, any>(
        (fileVersions || []).map((version: any) => [
          String(version.VersionLabel || ""),
          version,
        ]),
      );

      const currentLabel = String(fileInfo.OData__UIVersionString || "");
      const versionRowsSource: any[] =
        itemVersions && itemVersions.length
          ? itemVersions
          : [
              {
                ...fileInfo,
                VersionLabel: currentLabel,
                IsCurrentVersion: true,
                FieldValuesAsText: currentFieldValuesAsText,
              },
            ];

      // Map Sponsor lookup IDs -> titles (dropdown + current item).
      const sponsorIdTitleMap = new Map<number, string>();
      (drpdown.sponsor || []).forEach((option) => {
        const id = Number(option.code);
        if (id > 0 && option.name) {
          sponsorIdTitleMap.set(id, option.name);
        }
      });
      const currentSponsorId = Number(
        fileInfo[loanLibraryFields.sponsorNameId] ??
          fileInfo[loanLibraryFields.sponsorName]?.Id ??
          fileInfo[loanLibraryFields.sponsorName]?.ID ??
          rowData.sponsor?.id ??
          0,
      );
      const currentSponsorTitle =
        parseLookupValue(fileInfo[loanLibraryFields.sponsorName]) ||
        parseLookupValue(
          currentFieldValuesAsText?.[loanLibraryFields.sponsorName],
        ) ||
        rowData.sponsor?.sponsorTitle ||
        "";
      if (currentSponsorId > 0 && currentSponsorTitle) {
        sponsorIdTitleMap.set(currentSponsorId, currentSponsorTitle);
      }

      // Prefetch version payloads + FieldValuesAsText for historical versions.
      // Historical taxonomy fields often only contain WssIds; text endpoint restores labels.
      const versionPayloadByItemVersionId = new Map<number, Record<string, any>>();
      const versionTextByItemVersionId = new Map<number, Record<string, any>>();

      await Promise.all(
        versionRowsSource.map(async (versionItem: any) => {
          const itemVersionId = Number(
            versionItem.VersionId ?? versionItem.ID ?? versionItem.Id ?? 0,
          );
          const isCurrent =
            Boolean(versionItem.IsCurrentVersion) ||
            String(versionItem.VersionLabel || "") === currentLabel;

          if (isCurrent) {
            if (currentFieldValuesAsText) {
              versionTextByItemVersionId.set(
                itemVersionId || -1,
                currentFieldValuesAsText,
              );
            }
            return;
          }

          if (!itemVersionId || !itemId) {
            return;
          }

          const [payload, textValues] = await Promise.all([
            getListItemVersionPayload(itemId, itemVersionId),
            getListItemVersionFieldValuesAsText(itemId, itemVersionId),
          ]);

          if (payload) {
            versionPayloadByItemVersionId.set(itemVersionId, payload);
          }
          if (textValues) {
            versionTextByItemVersionId.set(itemVersionId, textValues);
          } else if (versionItem.FieldValuesAsText) {
            versionTextByItemVersionId.set(
              itemVersionId,
              versionItem.FieldValuesAsText,
            );
          }
        }),
      );

      // Collect taxonomy WssIds / TermGuids across versions for label resolution.
      const pendingWssIds: number[] = [];
      const pendingTermGuids: string[] = [];
      const collectFromPayload = (
        payload: Record<string, any> | null | undefined,
      ): void => {
        if (!payload) return;
        [
          managedMetadataFields.assetManagement,
          managedMetadataFields.legal,
          managedMetadataFields.servicing,
        ].forEach((fieldName) => {
          getTaxonomyFieldCandidates(payload, fieldName).forEach((candidate) => {
            pendingWssIds.push(...collectTaxonomyWssIds(candidate));
            pendingTermGuids.push(...collectTaxonomyTermGuids(candidate));
          });
        });
      };

      versionRowsSource.forEach((versionItem: any) => {
        const itemVersionId = Number(
          versionItem.VersionId ?? versionItem.ID ?? versionItem.Id ?? 0,
        );
        collectFromPayload(versionItem);
        collectFromPayload(versionPayloadByItemVersionId.get(itemVersionId));
        collectFromPayload(versionTextByItemVersionId.get(itemVersionId));
        collectFromPayload(versionItem.FieldValuesAsText);
      });
      collectFromPayload(fileInfo);
      collectFromPayload(currentFieldValuesAsText);

      const { wssIdLabelMap, termGuidLabelMap } = await resolveTaxonomyLabelMaps(
        pendingWssIds,
        pendingTermGuids,
      );

      const buildRowFromItemVersion = (
        versionItem: any,
        textValues?: Record<string, any> | null,
      ): IVersionHistoryRow => {
        const versionLabel = String(
          versionItem.VersionLabel || versionItem.versionLabel || "",
        );
        const fileVersion = fileVersionByLabel.get(versionLabel);
        const itemVersionId = Number(
          versionItem.VersionId ?? versionItem.ID ?? versionItem.Id ?? 0,
        );
        const enrichedPayload =
          versionPayloadByItemVersionId.get(itemVersionId) || versionItem;
        const fileVersionId =
          Number(fileVersion?.ID ?? fileVersion?.Id) || undefined;
        const isCurrent =
          Boolean(versionItem.IsCurrentVersion) ||
          versionLabel === currentLabel;

        const resolvedText =
          textValues ||
          versionTextByItemVersionId.get(itemVersionId) ||
          versionTextByItemVersionId.get(-1) ||
          enrichedPayload?.FieldValuesAsText ||
          versionItem.FieldValuesAsText ||
          (isCurrent ? currentFieldValuesAsText : null);

        const sponsorName =
          resolveSponsorFromVersion(
            enrichedPayload,
            resolvedText,
            sponsorIdTitleMap,
          ) || (isCurrent ? rowData.sponsor?.sponsorTitle || "" : "");

        const assetManagement = resolveTaxonomyFromVersion(
          enrichedPayload,
          managedMetadataFields.assetManagement,
          resolvedText,
          wssIdLabelMap,
          termGuidLabelMap,
        );
        const servicing = resolveTaxonomyFromVersion(
          enrichedPayload,
          managedMetadataFields.servicing,
          resolvedText,
          wssIdLabelMap,
          termGuidLabelMap,
        );
        const legal = resolveTaxonomyFromVersion(
          enrichedPayload,
          managedMetadataFields.legal,
          resolvedText,
          wssIdLabelMap,
          termGuidLabelMap,
        );

        const modifiedBy =
          parsePersonName(enrichedPayload.Editor) ||
          parsePersonName(fileVersion?.CreatedBy) ||
          parsePersonName(enrichedPayload.Author) ||
          "";

        const modifiedRaw =
          enrichedPayload.Modified ||
          enrichedPayload.Created ||
          fileVersion?.Created ||
          "";

        return {
          versionLabel,
          versionId: fileVersionId || itemVersionId || undefined,
          isCurrent,
          name:
            enrichedPayload.FileLeafRef ||
            resolvedText?.FileLeafRef ||
            fileInfo.FileLeafRef ||
            rowData.fileName ||
            "",
          sponsor: sponsorName,
          createdBy: parsePersonName(fileInfo.Author),
          createddate: formatDateTime(fileInfo.Created),
          modified: formatDateTime(modifiedRaw),
          modifiedBy,
          assetmanagement: assetManagement,
          servicing,
          legal,
          size: fileVersion?.Size
            ? `${(fileVersion.Size / 1024).toFixed(2)} KB`
            : "-",
          comments: isCurrent
            ? "Current Version"
            : fileVersion?.CheckInComment ||
              enrichedPayload.CheckInComment ||
              "-",
          url: isCurrent
            ? rowData.fileRef
            : buildHistoricalFileUrl(
                rowData.fileRef,
                webServerRelativeUrl,
                versionLabel,
              ) ||
              fileVersion?.Url ||
              fileVersion?.ServerRelativeUrl ||
              "",
        };
      };

      let history: IVersionHistoryRow[] = versionRowsSource.map(
        (versionItem: any) => buildRowFromItemVersion(versionItem),
      );

      // If item versions were empty, enrich historical rows from file versions.
      if ((!itemVersions || !itemVersions.length) && fileVersions?.length) {
        const currentRow = history[0];
        history = [
          currentRow,
          ...fileVersions
            .filter(
              (version: any) =>
                String(version.VersionLabel || "") !== currentLabel,
            )
            .map((version: any) => ({
              ...currentRow,
              versionLabel: String(version.VersionLabel || ""),
              versionId: Number(version.ID ?? version.Id) || undefined,
              isCurrent: false,
              modified: formatDateTime(version.Created),
              modifiedBy:
                parsePersonName(version.CreatedBy) || currentRow.modifiedBy,
              size: version.Size
                ? `${(version.Size / 1024).toFixed(2)} KB`
                : "-",
              comments: version.CheckInComment || "-",
              url:
                buildHistoricalFileUrl(
                  rowData.fileRef,
                  webServerRelativeUrl,
                  String(version.VersionLabel || ""),
                ) ||
                version.Url ||
                version.ServerRelativeUrl ||
                "",
            })),
        ];
      }

      // Ensure the live current version uses current FieldValuesAsText values.
      const currentIndex = history.findIndex(
        (row) => row.versionLabel === currentLabel || row.isCurrent,
      );
      if (currentIndex >= 0) {
        const currentFromFile = buildRowFromItemVersion(
          {
            ...fileInfo,
            VersionLabel: currentLabel || history[currentIndex].versionLabel,
            IsCurrentVersion: true,
          },
          currentFieldValuesAsText,
        );
        history[currentIndex] = {
          ...history[currentIndex],
          ...currentFromFile,
          isCurrent: true,
          comments: "Current Version",
        };
      } else if (currentLabel) {
        history.unshift(
          buildRowFromItemVersion(
            {
              ...fileInfo,
              VersionLabel: currentLabel,
              IsCurrentVersion: true,
            },
            currentFieldValuesAsText,
          ),
        );
      }

      // Deduplicate by version label (prefer rows with richer metadata).
      const deduped = new Map<string, IVersionHistoryRow>();
      history
        .sort(
          (a, b) => parseFloat(b.versionLabel) - parseFloat(a.versionLabel),
        )
        .forEach((row) => {
          const existing = deduped.get(row.versionLabel);
          if (!existing) {
            deduped.set(row.versionLabel, row);
            return;
          }

          deduped.set(row.versionLabel, {
            ...existing,
            ...row,
            sponsor: row.sponsor || existing.sponsor,
            assetmanagement: row.assetmanagement || existing.assetmanagement,
            servicing: row.servicing || existing.servicing,
            legal: row.legal || existing.legal,
          });
        });

      setVersionHistory(
        Array.from(deduped.values()).sort(
          (a, b) => parseFloat(b.versionLabel) - parseFloat(a.versionLabel),
        ),
      );
    } catch (error) {
      console.error("Version History Error:", error);
      toastFunc("error", "Error", "Failed to load version history");
      setShowVersionHistory(false);
      setVersionHistoryTarget(null);
    } finally {
      setVersionHistoryLoading(false);
    }
  };
  const openVersionAction = (
    action: "restore" | "delete",
    version: IVersionHistoryRow,
  ): void => {
    if (version.isCurrent) return;
    setVersionActionDialog({
      visible: true,
      action,
      version,
    });
  };

  // Close version restore/delete confirmation dialog.
  const closeVersionActionDialog = (): void => {
    if (versionActionLoading) return;
    setVersionActionDialog(EMPTY_VERSION_ACTION);
  };

  // Restore a historical version as the new current version.
  const restoreFileVersion = async (
    version: IVersionHistoryRow,
  ): Promise<void> => {
    if (!versionHistoryTarget?.fileRef || !version.versionLabel) return;

    try {
      setVersionActionLoading(true);
      await sp.web
        .getFileByServerRelativePath(versionHistoryTarget.fileRef)
        .versions.restoreByLabel(version.versionLabel);

      toastFunc(
        "success",
        "Success",
        `Version ${version.versionLabel} restored successfully`,
      );
      setVersionDetails(null);
      setVersionActionDialog(EMPTY_VERSION_ACTION);
      await getVersionHistory(versionHistoryTarget, true);
    } catch (error) {
      console.error("restoreFileVersion error:", error);
      toastFunc("error", "Error", "Failed to restore version");
    } finally {
      setVersionActionLoading(false);
    }
  };

  // Delete a historical version from the file version history.
  const deleteFileVersion = async (
    version: IVersionHistoryRow,
  ): Promise<void> => {
    if (!versionHistoryTarget?.fileRef || !version.versionLabel) return;

    try {
      setVersionActionLoading(true);
      const versions = sp.web.getFileByServerRelativePath(
        versionHistoryTarget.fileRef,
      ).versions;

      if (version.versionId) {
        await versions.deleteById(version.versionId);
      } else {
        await versions.deleteByLabel(version.versionLabel);
      }

      toastFunc(
        "success",
        "Success",
        `Version ${version.versionLabel} deleted successfully`,
      );
      setVersionDetails(null);
      setVersionActionDialog(EMPTY_VERSION_ACTION);
      await getVersionHistory(versionHistoryTarget, true);
    } catch (error) {
      console.error("deleteFileVersion error:", error);
      toastFunc("error", "Error", "Failed to delete version");
    } finally {
      setVersionActionLoading(false);
    }
  };

  // Confirm and run the selected version restore or delete action.
  const confirmVersionAction = async (): Promise<void> => {
    if (!versionActionDialog.version || !versionActionDialog.action) return;

    if (versionActionDialog.action === "restore") {
      await restoreFileVersion(versionActionDialog.version);
      return;
    }

    await deleteFileVersion(versionActionDialog.version);
  };

  const closeVersionHistoryDialog = (): void => {
    if (versionHistoryLoading || versionActionLoading) return;
    setShowVersionHistory(false);
    setVersionHistoryTarget(null);
    setVersionHistory([]);
    setVersionDetails(null);
    setVersionActionDialog(EMPTY_VERSION_ACTION);
  };

  const openVersionActionFromDetails = (
    action: "restore" | "delete",
  ): void => {
    if (!versionDetails || versionDetails.isCurrent) return;
    const selectedVersion = versionDetails;
    openVersionAction(action, selectedVersion);
  };

  // Open the exact saved file content for the selected version.
  const openFileVersionContent = (version: IVersionHistoryRow): void => {
    const versionUrl =
      version.url ||
      (version.isCurrent ? versionHistoryTarget?.fileRef || "" : "");

    if (!versionUrl) {
      toastFunc("warn", "Version unavailable", "This version cannot be opened");
      return;
    }

    const absoluteUrl = /^https?:\/\//i.test(versionUrl)
      ? versionUrl
      : new URL(versionUrl, window.location.origin).href;
    const fileName =
      version.name || versionHistoryTarget?.fileName || "";
    const isOfficeDocument = isOfficeDocumentName(fileName);

    if (isOfficeDocument && !version.isCurrent) {
      const officeProtocol = getOfficeDesktopProtocol(fileName);
      window.location.href = `${officeProtocol}:ofv|u|${absoluteUrl}`;
      return;
    }

    const browserUrl =
      isOfficeDocument && version.isCurrent
        ? `${absoluteUrl}${absoluteUrl.includes("?") ? "&" : "?"}web=1`
        : absoluteUrl;
    window.open(browserUrl, "_blank", "noopener,noreferrer");
  };

  // Render version history action links (View / Open / Restore / Delete).
  const versionActionsTemplate = (
    row: IVersionHistoryRow,
  ): React.ReactElement => {
    return (
      <div className={styles.versionActions}>
        <button
          type="button"
          className={styles.versionActionView}
          onClick={() => setVersionDetails(row)}
          disabled={versionActionLoading}
        >
          <i className="pi pi-eye" />
          View
        </button>
        {!row.isCurrent && (
          <>
            <span className={styles.versionActionDivider}>|</span>
            <button
              type="button"
              className={styles.versionActionRestore}
              onClick={() => openVersionAction("restore", row)}
              disabled={versionActionLoading}
            >
              <i className="pi pi-replay" />
              Restore
            </button>
            <span className={styles.versionActionDivider}>|</span>
            <button
              type="button"
              className={styles.versionActionDelete}
              onClick={() => openVersionAction("delete", row)}
              disabled={versionActionLoading}
            >
              <i className="pi pi-trash" />
              Delete
            </button>
          </>
        )}
      </div>
    );
  };

  const versionCellTemplate = (value?: string): React.ReactElement => {
    const empty = !value || value === "-";

    return (
      <span className={empty ? styles.versionCellEmpty : styles.versionCellText}>
        {value || "—"}
      </span>
    );
  };

  // Collect non-empty taxonomy tags across AM / Legal / Servicing.
  const getVersionTagValues = (row: IVersionHistoryRow): string[] => {
    const labels = [
      ...toTaxonomyLabels(row.assetmanagement),
      ...toTaxonomyLabels(row.legal),
      ...toTaxonomyLabels(row.servicing),
    ];
    return Array.from(new Set(labels));
  };

  // Render one version history card row.
  const versionHistoryRowTemplate = (
    row: IVersionHistoryRow,
  ): React.ReactElement => {
    const tags = getVersionTagValues(row);
    const showComments =
      !!row.comments &&
      row.comments !== "-" &&
      !(row.isCurrent && row.comments === "Current Version");

    return (
      <div
        key={row.versionLabel}
        className={styles.versionTimelineItem}
      >
        <div className={styles.versionTimelineRail}>
          <span
            className={`${styles.versionTimelineMarker} ${
              row.isCurrent ? styles.versionTimelineMarkerCurrent : ""
            }`}
          >
            <i className={row.isCurrent ? "pi pi-check" : "pi pi-history"} />
          </span>
          <span className={styles.versionTimelineLine} />
        </div>

        <div
          className={`${styles.versionCard} ${
            row.isCurrent ? styles.versionCardCurrent : ""
          }`}
        >
          <div className={styles.versionCardHeader}>
            <div className={styles.versionCardHeading}>
              <div className={styles.versionLabelCell}>
                <span className={styles.versionNumber}>
                  Version {row.versionLabel}
                </span>
                {row.isCurrent ? (
                  <span className={styles.versionCurrentBadge}>Current</span>
                ) : null}
              </div>
              <div className={styles.versionCardActivity}>
                <span>
                  <i className="pi pi-clock" />
                  {row.modified || "—"}
                </span>
                <span>
                  <i className="pi pi-user" />
                  {row.modifiedBy || "—"}
                </span>
              </div>
            </div>
            {versionActionsTemplate(row)}
          </div>

          <div className={styles.versionCardFacts}>
            {row.sponsor && row.sponsor !== "-" && (
              <span className={styles.versionFact}>
                <i className="pi pi-building" />
                {row.sponsor}
              </span>
            )}
            {row.size && row.size !== "-" && (
              <span className={styles.versionFact}>
                <i className="pi pi-file" />
                {row.size}
              </span>
            )}
          </div>

          {tags.length > 0 ? (
            <div className={styles.versionCardTags}>
              {tags.map((tag) => (
                <span
                  key={`${row.versionLabel}-${tag}`}
                  className={styles.metadataTagPill}
                >
                  {tag}
                </span>
              ))}
            </div>
          ) : null}

          {showComments ? (
            <div className={styles.versionCardComment}>
              <i className="pi pi-comment" />
              <span>{row.comments}</span>
            </div>
          ) : null}
        </div>
      </div>
    );
  };

  // Render action menu trigger button for each table row.
  const actionTemplate = (rowData: any): React.ReactElement => (
    <Button
      icon="pi pi-ellipsis-h"
      text
      rounded
      className={styles.rowActionBtn}
      onClick={(e: any) => {
        e.stopPropagation();
        selectedRowRef.current = rowData;
        getRowMenuItems(rowData);
        menuRef.current?.toggle(e);
      }}
    />
  );
  //folder file icon
  const getFileIcon = (fileName: string): string => {
    const extension = fileName.split(".").pop()?.toLowerCase();

    switch (extension) {
      case "doc":
      case "docx":
        return "pi pi-file-word";

      case "xls":
      case "xlsx":
        return "pi pi-file-excel";

      case "ppt":
      case "pptx":
        return "pi pi-file";

      case "pdf":
        return "pi pi-file-pdf";

      case "jpg":
      case "jpeg":
      case "png":
      case "gif":
      case "bmp":
      case "webp":
        return "pi pi-image";

      case "zip":
      case "rar":
        return "pi pi-box";

      default:
        return "pi pi-file";
    }
  };
  //sub folder create
  const createsubFolder = async (): Promise<void> => {
    try {
      if (!folderName?.trim()) {
        toastFunc("warn", "Validation", "Please enter folder name");
        return;
      }
      const trimmedName = folderName.trim();
      const currentFolderItems =
        folderCacheRef.current.get(currentFolder) || displayItems;
      const folderAlreadyExists = currentFolderItems.some(
        (item) =>
          item.folderType === 1 &&
          String(item.fileName || "").localeCompare(trimmedName, undefined, {
            sensitivity: "base",
          }) === 0,
      );

      if (folderAlreadyExists) {
        toastFunc(
          "error",
          "Folder already exists",
          `A folder named "${trimmedName}" already exists in this location. Please choose a different name.`,
        );
        return;
      }

      setLoader(true);
      const folderPath = `${currentFolder}/${trimmedName}`;

      // Resolve sponsor value from the root folder record.
      const rootLoanFolder = mainFolders.find((f) =>
        currentFolder.includes(f.fileRef),
      );
      const sponsorId = rootLoanFolder?.sponsor?.id ?? null;

      await sp.web.folders.addUsingPath(folderPath);

      // Save sponsor lookup value on the new subfolder item.
      if (sponsorId) {
        const folderItem: any = await sp.web
          .getFolderByServerRelativePath(folderPath)
          .listItemAllFields();

        await sp.web.lists
          .getByTitle(listNames.loan)
          .items.getById(folderItem.Id)
          .update(buildSponsorLookupUpdatePayload(sponsorId));
      }

      toastFunc("success", "Success", "Folder created successfully");
      setFolderName("");
      setShowCreateFolder(false);

      const updatedItems = await fetchFolderContents(currentFolder);
      folderCacheRef.current.set(currentFolder, updatedItems);
      invalidateLoanTreeCacheForPath(currentFolder);
      await refreshCurrentView();
    } catch (error) {
      console.error("Create Folder Error:", error);
      const rawErrorMessage = [
        String((error as any)?.message || ""),
        String((error as any)?.data?.responseBody || ""),
        String(error || ""),
      ].join(" ");
      const isDuplicateFolder =
        /already exists|same name|duplicate/i.test(rawErrorMessage) ||
        rawErrorMessage.includes("0x80070050") ||
        rawErrorMessage.includes("-2130575257");

      toastFunc(
        "error",
        isDuplicateFolder ? "Folder already exists" : "Unable to create folder",
        isDuplicateFolder
          ? `A folder named "${folderName.trim()}" already exists in this location. Please choose a different name.`
          : "The folder could not be created. Please check the folder name and try again.",
      );
    } finally {
      setLoader(false);
    }
  };

  const getSharePointErrorMessage = (error: any): string => {
    const responseBody = String(
      error?.data?.responseBody ||
        error?.responseBody ||
        error?.message ||
        "",
    );

    try {
      const parsed =
        typeof error?.data?.responseBody === "object"
          ? error.data.responseBody
          : JSON.parse(responseBody);
      const odataMessage =
        parsed?.["odata.error"]?.message?.value ||
        parsed?.error?.message?.value ||
        parsed?.error?.message;
      if (odataMessage) {
        return String(odataMessage);
      }
    } catch {
      // Fall through to regex parsing for non-JSON payloads.
    }

    const responseMessageMatch = responseBody.match(
      /"value"\s*:\s*"([^"]+)"/i,
    );
    return (
      responseMessageMatch?.[1] ||
      String(error?.message || "").trim() ||
      "The folder request could not be submitted. Please try again."
    );
  };

  const resolveListFieldInternalName = async (
    list: any,
    fieldName: string,
  ): Promise<string> => {
    try {
      const field: any = await list.fields
        .getByInternalNameOrTitle(fieldName)
        .select("InternalName")();
      return String(field.InternalName || fieldName);
    } catch {
      return fieldName;
    }
  };

  const submitFolderRequest = async (): Promise<void> => {
    const requestedFolderName = folderName.trim();
    if (!requestedFolderName) {
      toastFunc("warn", "Validation", "Please enter folder name");
      return;
    }

    const folderAlreadyExists = displayItems.some(
      (item) =>
        item.folderType === 1 &&
        item.fileName.localeCompare(requestedFolderName, undefined, {
          sensitivity: "base",
        }) === 0,
    );
    if (folderAlreadyExists) {
      toastFunc(
        "error",
        "Folder already exists",
        `A folder named "${requestedFolderName}" already exists in this location.`,
      );
      return;
    }

    try {
      setRequestFolderLoading(true);

      const requestList = sp.web.lists.getByTitle(listNames.loanFolderRequest);

      // Confirm the list exists on the current site before saving.
      try {
        await requestList.select("Id", "Title")();
      } catch {
        throw new Error(
          `List '${listNames.loanFolderRequest}' was not found on this site. Create it and try again.`,
        );
      }

      const loanList = sp.web.lists.getByTitle(listNames.loan);
      const parentFolderUrl = currentFolder
        ? normalizeSharePointPath(currentFolder)
        : normalizeSharePointPath(
            (await loanList.rootFolder.select("ServerRelativeUrl")())
              .ServerRelativeUrl,
          );

      if (parentFolderUrl.length > 255) {
        throw new Error(
          "Parent folder path exceeds 255 characters for ParentFolderUrl.",
        );
      }

      const loanFolder =
        mainFolders
          .filter(
            (loan) =>
              !!loan.fileRef &&
              (!!currentFolder
                ? isFolderDescendantOrSelf(loan.fileRef, currentFolder)
                : false),
          )
          .sort(
            (left, right) =>
              normalizeSharePointPath(right.fileRef).length -
              normalizeSharePointPath(left.fileRef).length,
          )[0] || null;

      const loanNumber = loanFolder?.fileName || requestedFolderName;
      const loanItemId = Number(loanFolder?.id || 0);

      const [
        titleField,
        parentFolderUrlField,
        loanItemIdField,
        loanNumberField,
        requestStatusField,
      ] = await Promise.all([
        resolveListFieldInternalName(requestList, "Title"),
        resolveListFieldInternalName(requestList, "ParentFolderUrl"),
        resolveListFieldInternalName(requestList, "LoanItemId"),
        resolveListFieldInternalName(requestList, "LoanNumber"),
        resolveListFieldInternalName(requestList, "RequestStatus"),
      ]);

      // Soft duplicate check — never block save if the filter query fails.
      try {
        const existingRequests: any[] = await requestList.items
          .select("Id", titleField, parentFolderUrlField, requestStatusField)
          .filter(
            `${titleField} eq '${escapeODataString(requestedFolderName)}' and ` +
              `${requestStatusField} eq 'Pending'`,
          )
          .top(5000)();

        const duplicateRequest = existingRequests.some((request) =>
          pathsEqual(
            getServerRelativePathFromUrl(
              getSharePointUrlValue(request[parentFolderUrlField]),
            ),
            parentFolderUrl,
          ),
        );

        if (duplicateRequest) {
          toastFunc(
            "error",
            "Request already exists",
            `A pending request for "${requestedFolderName}" already exists in this location.`,
          );
          return;
        }
      } catch (duplicateCheckError) {
        console.warn(
          "LoanFolderRequest duplicate check skipped:",
          duplicateCheckError,
        );
      }

      const requestPayload: Record<string, any> = {
        [titleField]: requestedFolderName,
        [parentFolderUrlField]: parentFolderUrl,
        [loanItemIdField]: loanItemId,
        [loanNumberField]: loanNumber,
        [requestStatusField]: "Pending",
      };

      const addResult: any = await requestList.items.add(requestPayload);
      const createdItemId = Number(
        addResult?.data?.Id || addResult?.data?.ID || addResult?.Id || 0,
      );

      if (!createdItemId) {
        throw new Error(
          "SharePoint did not return a created item ID. The request may not have been saved.",
        );
      }

      toastFunc(
        "success",
        "Folder requested",
        `"${requestedFolderName}" was saved to ${listNames.loanFolderRequest} (ID ${createdItemId}).`,
      );
      setFolderName("");
      isRequestFolderDialogRef.current = false;
      setShowCreateFolder(false);
      setIsRequestFolderDialog(false);

      // After a successful request, return to the main loan folder list.
      await navigateToFolder("");
    } catch (error) {
      console.error("Submit folder request error:", error);
      toastFunc("error", "Request failed", getSharePointErrorMessage(error));
    } finally {
      setRequestFolderLoading(false);
    }
  };

  // Reset file upload dialog fields.
  const resetFileUploadState = (): void => {
    preserveFileUploadStateOnOpenRef.current = false;
    setUploadFiles([]);
    setUploadDragOver(false);
    setUploadTagTermIds([]);
    setFileUploading(false);
    setUploadProgress({
      completed: 0,
      total: 0,
      currentFileName: "",
      successCount: 0,
      failCount: 0,
      phase: "uploading",
    });
    if (fileUploadInputRef.current) {
      fileUploadInputRef.current.value = "";
    }
  };

  const closeFileUploadDialog = (): void => {
    if (fileUploading) return;
    resetFileUploadState();
    setShowFileUpload(false);
  };

  const closeEditTagDialog = (): void => {
    if (editTagSaving) return;
    setShowEditTagDialog(false);
    setEditTagTarget(null);
    setEditTagCategory(null);
    setEditTagLoading(false);
    setEditTagTermIds({ ...EMPTY_EDIT_TAG_TERM_IDS });
  };

  const openEditTagDialog = async (row: ILoanRecord): Promise<void> => {
    if (!row?.id) return;

    const category = detectTaxonomyCategoryFromPath(currentFolder);
    if (!category) {
      toastFunc(
        "warn",
        "Warning",
        "Tags can only be edited inside Asset Management, Legal, or Servicing folders",
      );
      return;
    }

    setEditTagTarget(row);
    setEditTagCategory(category);
    setShowEditTagDialog(true);
    setEditTagLoading(true);
    setEditTagTermIds({ ...EMPTY_EDIT_TAG_TERM_IDS });

    try {
      const { trees, fields } = await loadManagedMetadataTagTreesCached();
      setEditTagTrees(trees);
      setEditTagFields(fields);

      const fieldName = fields[category];
      const item: any = await sp.web.lists
        .getByTitle(listNames.loan)
        .items.getById(row.id)
        .select(fieldName)
        .get();

      setEditTagTermIds({
        ...EMPTY_EDIT_TAG_TERM_IDS,
        [category]: resolveTermIdsForCategory(
          item[fieldName],
          row[TAXONOMY_CATEGORY_ROW_FIELD[category]] || "",
          trees[category] || [],
        ),
      });
    } catch (error) {
      console.error("openEditTagDialog error:", error);
      toastFunc("error", "Error", "Failed to load tags for this file");
      closeEditTagDialog();
    } finally {
      setEditTagLoading(false);
    }
  };

  const saveEditTags = async (): Promise<void> => {
    if (
      !editTagTarget?.id ||
      !editTagCategory ||
      editTagSaving ||
      editTagLoading
    ) {
      return;
    }

    setEditTagSaving(true);
    try {
      const tags = buildTaxonomyTags(
        editTagTermIds[editTagCategory] || [],
        editTagCategory,
        editTagTrees,
        editTagFields,
      );
      await updateTaxonomyFieldOnListItem(
        editTagTarget.id,
        editTagFields[editTagCategory],
        tags,
      );

      const rowField = TAXONOMY_CATEGORY_ROW_FIELD[editTagCategory];
      const tagUpdates = {
        [rowField]: buildTaxonomyDisplayValue(
          editTagTermIds[editTagCategory] || [],
          editTagCategory,
          editTagTrees,
        ),
      };

      const levelPath = currentFolder;
      const cachedItems =
        folderCacheRef.current.get(levelPath) ?? displayItems;
      const updatedItems = cachedItems.map((item) =>
        item.id === editTagTarget.id ? { ...item, ...tagUpdates } : item,
      );

      folderCacheRef.current.set(levelPath, updatedItems);
      setDisplayItems(applyFilter(updatedItems, filter));

      toastFunc("success", "Success", "Tags updated successfully");
      closeEditTagDialog();
    } catch (error) {
      console.error("saveEditTags error:", error);
      toastFunc(
        "error",
        "Error",
        error instanceof Error ? error.message : "Failed to update tags",
      );
    } finally {
      setEditTagSaving(false);
    }
  };

  // Load taxonomy trees when the upload dialog opens.
  useEffect(() => {
    if (!showFileUpload) return;
    if (preserveFileUploadStateOnOpenRef.current) {
      preserveFileUploadStateOnOpenRef.current = false;
      return;
    }

    setUploadTagCategory(detectTaxonomyCategoryFromPath(currentFolder));
    setUploadTagTermIds([]);
    setUploadFiles([]);
    setUploadDragOver(false);

    let cancelled = false;
    const loadTags = async (): Promise<void> => {
      try {
        setUploadTagsLoading(true);
        const { trees, fields } = await loadManagedMetadataTagTreesCached();
        if (cancelled) return;
        setUploadTagTrees(trees);
        setUploadTagFields(fields);
      } catch (error) {
        console.error("Failed to load taxonomy tags:", error);
        if (!cancelled) {
          toastFunc("error", "Error", "Failed to load managed metadata tags");
        }
      } finally {
        if (!cancelled) setUploadTagsLoading(false);
      }
    };

    void loadTags();
    return () => {
      cancelled = true;
    };
  }, [showFileUpload, currentFolder]);

  const addFilesToUpload = (files: File[]): void => {
    if (!files.length) return;

    const nextFiles = [...uploadFiles];
    let duplicateCount = 0;

    for (const file of files) {
      const isDuplicate = nextFiles.some(
        (selectedFile) =>
          selectedFile.name.toLowerCase() === file.name.toLowerCase() &&
          selectedFile.size === file.size &&
          selectedFile.lastModified === file.lastModified,
      );

      if (isDuplicate) {
        duplicateCount += 1;
        continue;
      }

      if (nextFiles.length >= 5) break;
      nextFiles.push(file);
    }

    setUploadFiles(nextFiles);

    if (files.length - duplicateCount > 5 - uploadFiles.length) {
      toastFunc(
        "warn",
        "Maximum files reached",
        "You can upload a maximum of 5 files at a time.",
      );
    } else if (duplicateCount) {
      toastFunc(
        "info",
        "Duplicate file ignored",
        `${duplicateCount} already selected ${
          duplicateCount === 1 ? "file was" : "files were"
        } ignored.`,
      );
    }
  };

  const onUploadDrop = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    e.stopPropagation();
    setUploadDragOver(false);
    const files = e.dataTransfer.files;
    if (!files?.length) return;
    addFilesToUpload(Array.from(files));
  };

  const onUploadFileInputChange = (
    e: React.ChangeEvent<HTMLInputElement>,
  ): void => {
    const files = e.target.files;
    if (!files?.length) return;
    addFilesToUpload(Array.from(files));
    e.target.value = "";
  };

  // Upload up to five files with the same sponsor and managed metadata tags.
  const confirmFileUpload = async (): Promise<void> => {
    if (!uploadFiles.length) {
      toastFunc("warn", "Validation", "Please select at least one file");
      return;
    }
    if (!currentFolder) {
      toastFunc("error", "Error", "No destination folder selected");
      return;
    }

    // Block re-upload of files that already exist in the current folder.
    const folderItems =
      folderCacheRef.current.get(currentFolder) || displayItems;
    const existingFileNames = new Set(
      folderItems
        .filter(
          (item) =>
            item.folderType !== 1 &&
            !item.isHyperlink &&
            !!String(item.fileName || "").trim(),
        )
        .map((item) => String(item.fileName).trim().toLowerCase()),
    );
    const alreadyUploadedFiles = uploadFiles.filter((file) =>
      existingFileNames.has(file.name.trim().toLowerCase()),
    );

    if (alreadyUploadedFiles.length) {
      toastFunc(
        "error",
        "File already uploaded",
        alreadyUploadedFiles.length === 1
          ? `"${alreadyUploadedFiles[0].name}" is already uploaded to this folder.`
          : `${alreadyUploadedFiles
              .map((file) => `"${file.name}"`)
              .join(", ")} ${
              alreadyUploadedFiles.length === 1 ? "is" : "are"
            } already uploaded to this folder.`,
      );
      return;
    }

    try {
      setShowFileUpload(false);
      setFileUploading(true);
      setUploadProgress({
        completed: 0,
        total: uploadFiles.length,
        currentFileName: uploadFiles[0]?.name || "",
        successCount: 0,
        failCount: 0,
        phase: "uploading",
      });
      const rootLoanFolder = mainFolders.find((f) =>
        currentFolder.includes(f.fileRef),
      );
      const sponsorId = rootLoanFolder?.sponsor?.id ?? null;
      const taxonomyTags =
        uploadTagCategory && uploadTagTermIds.length
          ? buildTaxonomyTags(
          uploadTagTermIds,
          uploadTagCategory,
          uploadTagTrees,
          uploadTagFields,
            )
          : [];
      const failedFiles: Array<{ file: File; message: string }> = [];
      let uploadedCount = 0;

      for (let fileIndex = 0; fileIndex < uploadFiles.length; fileIndex += 1) {
        const file = uploadFiles[fileIndex];
        setUploadProgress((currentProgress) => ({
          ...currentProgress,
          currentFileName: file.name,
        }));

        try {
          // overwrite=false so SharePoint rejects an existing same-name file.
          const result = await sp.web
            .getFolderByServerRelativePath(currentFolder)
            .files.add(file.name, file, false);

          const fileItem: any = await result.file.listItemAllFields();
          const itemId = Number(fileItem.Id ?? fileItem.ID);
          if (!itemId) {
            throw new Error(
              `Could not resolve the list item ID for "${file.name}"`,
            );
          }

          if (sponsorId && sponsorId > 0) {
            await sp.web.lists
              .getByTitle(listNames.loan)
              .items.getById(itemId)
              .update(buildSponsorLookupUpdatePayload(sponsorId));
          }

          if (taxonomyTags.length) {
            await applyTaxonomyToListItem(itemId, taxonomyTags);
          }

          uploadedCount += 1;
        } catch (error) {
          console.error(`File upload failed for ${file.name}:`, error);
          const rawMessage = [
            String((error as any)?.message || ""),
            String((error as any)?.data?.responseBody || ""),
            String(error || ""),
          ].join(" ");
          const isAlreadyUploaded =
            /already exists|same name|duplicate|1837|0x800700b7|-2130575257/i.test(
              rawMessage,
            );

          failedFiles.push({
            file,
            message: isAlreadyUploaded
              ? `"${file.name}" is already uploaded to this folder.`
              : error instanceof Error
                ? error.message
                : "Upload failed",
          });
        } finally {
          setUploadProgress((currentProgress) => ({
            ...currentProgress,
            completed: fileIndex + 1,
            successCount: uploadedCount,
            failCount: failedFiles.length,
          }));
        }
      }

      setUploadProgress((currentProgress) => ({
        ...currentProgress,
        completed: uploadFiles.length,
        successCount: uploadedCount,
        failCount: failedFiles.length,
        phase: "complete",
      }));
      await new Promise((resolve) => setTimeout(resolve, 800));

      if (uploadedCount) {
        const updatedItems = await fetchFolderContents(currentFolder);
        folderCacheRef.current.set(currentFolder, updatedItems);
        invalidateLoanTreeCacheForPath(currentFolder);
        await refreshCurrentView();
      }

      if (failedFiles.length) {
        setUploadFiles(failedFiles.map(({ file }) => file));
        preserveFileUploadStateOnOpenRef.current = true;
        setShowFileUpload(true);
        const allAlreadyUploaded = failedFiles.every((item) =>
          /already uploaded/i.test(item.message),
        );
        toastFunc(
          "error",
          allAlreadyUploaded
            ? "File already uploaded"
            : uploadedCount
              ? "Upload partially completed"
              : "Upload failed",
          allAlreadyUploaded
            ? failedFiles.map(({ message }) => message).join(" ")
            : uploadedCount
              ? `${uploadedCount} ${
                  uploadedCount === 1 ? "file was" : "files were"
                } uploaded. ${failedFiles.length} failed and remain selected.`
              : `Could not upload: ${failedFiles
                  .map(({ file }) => file.name)
                  .join(", ")}`,
        );
      } else {
        toastFunc(
          "success",
          "Success",
          `${uploadedCount} ${
            uploadedCount === 1 ? "file was" : "files were"
          } uploaded successfully.`,
        );
        resetFileUploadState();
        setShowFileUpload(false);
      }
    } catch (error) {
      console.error("File upload failed:", error);
      preserveFileUploadStateOnOpenRef.current = true;
      setShowFileUpload(true);
      toastFunc(
        "error",
        "Error",
        error instanceof Error ? error.message : "Failed to upload file",
      );
    } finally {
      setFileUploading(false);
    }
  };

  // Close delete confirmation dialog.
  const closeDeleteDialog = (): void => {
    if (deleteLoading) return;
    setshowDeleteDialog(false);
    setDeleteTarget(null);
  };

  const getSourceFileMappingIds = async (
    sourceFileItemId: number,
  ): Promise<number[]> => {
    const mappingList = sp.web.lists.getByTitle(listNames.loanFileMapping);
    let paged = await mappingList.items
      .select("Id")
      .filter(`SourceFileItemId eq ${sourceFileItemId}`)
      .top(5000)
      .getPaged();
    const mappingIds = paged.results.map((mapping: any) =>
      Number(mapping.Id),
    );

    while (paged.hasNext) {
      paged = await paged.getNext();
      mappingIds.push(
        ...paged.results.map((mapping: any) => Number(mapping.Id)),
      );
    }

    return mappingIds.filter((mappingId) => mappingId > 0);
  };

  const deleteSourceFileMappings = async (
    mappingIds: number[],
  ): Promise<{ deleted: number; deactivated: number; failed: number }> => {
    const mappingList = sp.web.lists.getByTitle(listNames.loanFileMapping);
    let deleted = 0;
    let deactivated = 0;
    let failed = 0;

    for (const mappingId of mappingIds) {
      try {
        await mappingList.items.getById(mappingId).delete();
        deleted += 1;
      } catch (deleteError) {
        console.error(
          `Delete LoanFileMapping item ${mappingId} failed:`,
          deleteError,
        );
        try {
          await mappingList.items.getById(mappingId).update({
            IsActive: false,
          });
          deactivated += 1;
        } catch (deactivateError) {
          console.error(
            `Deactivate LoanFileMapping item ${mappingId} failed:`,
            deactivateError,
          );
          failed += 1;
        }
      }
    }

    return { deleted, deactivated, failed };
  };

  // Delete a file or folder from the SharePoint library (recycle bin).
  const confirmDeleteItem = async (): Promise<void> => {
    const row = deleteTarget || selectedRowRef.current;
    if (!row?.fileRef) {
      toastFunc("error", "Error", "Invalid item selected");
      return;
    }

    const isFolder = row.folderType === 1;
    const itemName = row.fileName || (isFolder ? "folder" : "file");
    const itemPath = row.fileRef;
    const sourceFileItemId = Number(row.id || 0);

    // Clear selection before the async call to avoid stale ref race conditions.
    selectedRowRef.current = null;

    try {
      setDeleteLoading(true);
      const mappingIds =
        !isFolder && sourceFileItemId > 0
          ? await getSourceFileMappingIds(sourceFileItemId)
          : [];

      if (isFolder) {
        await sp.web.getFolderByServerRelativePath(itemPath).recycle();
      } else {
        await sp.web.getFileByServerRelativePath(itemPath).recycle();
      }

      const mappingCleanup = mappingIds.length
        ? await deleteSourceFileMappings(mappingIds)
        : { deleted: 0, deactivated: 0, failed: 0 };

      await logActivity({
        title: itemName,
        action: "Delete",
        fromPath: itemPath,
        toPath: "",
        result: mappingCleanup.failed ? "Failed" : "Success",
      });

      if (mappingCleanup.failed) {
        toastFunc(
          "error",
          "File deleted; mapping cleanup incomplete",
          `${mappingCleanup.failed} mapping ${
            mappingCleanup.failed === 1 ? "record could" : "records could"
          } not be removed or deactivated.`,
        );
      } else {
        const cleanedMappingCount =
          mappingCleanup.deleted + mappingCleanup.deactivated;
        toastFunc(
          "success",
          "Success",
          `"${itemName}" moved to the recycle bin${
            cleanedMappingCount
              ? ` and ${cleanedMappingCount} mapping ${
                  cleanedMappingCount === 1 ? "record was" : "records were"
                } removed`
              : ""
          }`,
        );
      }

      setshowDeleteDialog(false);
      setDeleteTarget(null);

      if (mappingIds.length) {
        clearFolderCaches();
      }

      // Refresh the current folder / dashboard list.
      if (!currentFolder) {
        await getLoanData(drpdown.sponsor, false);
      } else {
        folderCacheRef.current.delete(currentFolder);
        invalidateLoanTreeCacheForPath(currentFolder);
        const items = await fetchFolderContents(currentFolder);
        folderCacheRef.current.set(currentFolder, items);
        await refreshCurrentView();
      }
    } catch (error) {
      console.error("confirmDeleteItem error:", error);
      await logActivity({
        title: itemName,
        action: "Delete",
        fromPath: itemPath,
        toPath: "",
        result: "Failed",
      });
      toastFunc(
        "error",
        "Error",
        `Failed to delete ${isFolder ? "folder" : "file"}`,
      );
    } finally {
      setDeleteLoading(false);
    }
  };

  // share functionality — embed SharePoint dialog in-page (no separate browser window)
  const closeShareDialog = (): void => {
    setShowShareDialog(false);
    setShareDialogUrl("");
    setShareItemName("");
  };

  const openNativeShareDialog = async (rowData: ILoanRecord): Promise<void> => {
    try {
      const library = await sp.web.lists
        .getByTitle(listNames.loan)
        .select("Id")();

      const shareUrl =
        `${props.context.pageContext.web.absoluteUrl}/_layouts/15/sharedialog.aspx` +
        `?listId=${encodeURIComponent(library.Id)}` +
        `&listItemId=${rowData.id}` +
        `&itemName=${encodeURIComponent(rowData.fileName)}` +
        `&clientId=SPList` +
        `&ma=0` +
        `&IsDlg=1`;

      setShareItemName(rowData.fileName);
      setShareDialogUrl(shareUrl);
      setShowShareDialog(true);
    } catch (error) {
      console.error("Open Share Dialog Error:", error);
      toastFunc("error", "Error", "Failed to open share dialog");
    }
  };
  // Hyperlink a source file into another loan under the same sponsor.
  const closeHyperlinkDialog = (allowWhileLoading = false): void => {
    if (hyperlinkLoading && !allowWhileLoading) return;
    setShowHyperlinkDialog(false);
    setHyperlinkTarget(null);
    setHyperlinkSourceLoan(null);
    setHyperlinkLoanOptions([]);
    setSelectedHyperlinkLoans([]);
  };

  const openHyperlinkDialog = (file: ILoanRecord): void => {
    const sourcePath = normalizeSharePointPath(file.fileRef);
    const sourceLoan =
      mainFolders
        .filter(
          (loan) =>
            !!loan.fileRef &&
            isFolderDescendantOrSelf(loan.fileRef, sourcePath),
        )
        .sort(
          (left, right) =>
            normalizeSharePointPath(right.fileRef).length -
            normalizeSharePointPath(left.fileRef).length,
        )[0] || null;
    const sponsorId = Number(
      sourceLoan?.sponsor?.id ?? file.sponsor?.id ?? 0,
    );

    if (!sourceLoan || !sponsorId) {
      toastFunc(
        "warn",
        "Parent sponsor not found",
        "The selected file is not associated with a parent sponsor.",
      );
      return;
    }

    const associatedLoans = mainFolders
      .filter(
        (loan) =>
          Number(loan.sponsor?.id || 0) === sponsorId &&
          !pathsEqual(loan.fileRef, sourceLoan.fileRef),
      )
      .sort((left, right) =>
        String(left.fileName || left.name || "").localeCompare(
          String(right.fileName || right.name || ""),
          undefined,
          { numeric: true, sensitivity: "base" },
        ),
      );

    setHyperlinkTarget(file);
    setHyperlinkSourceLoan(sourceLoan);
    setHyperlinkLoanOptions(associatedLoans);
    setSelectedHyperlinkLoans([]);
    setShowHyperlinkDialog(true);
  };

  const saveFileHyperlink = async (): Promise<void> => {
    if (
      !hyperlinkTarget?.id ||
      !hyperlinkTarget.fileRef ||
      !hyperlinkSourceLoan?.fileRef ||
      !selectedHyperlinkLoans.length
    ) {
      toastFunc(
        "warn",
        "Validation",
        "Select at least one destination loan number.",
      );
      return;
    }

    const sourceFilePath = normalizeSharePointPath(hyperlinkTarget.fileRef);
    const sourceLoanPath = normalizeSharePointPath(hyperlinkSourceLoan.fileRef);
    const sourceParentPath = getParentFolderPath(sourceFilePath);
    const relativeFolderPath = sourceParentPath
      .substring(sourceLoanPath.length)
      .replace(/^\/+|\/+$/g, "");
    const sourceFileUrl = toAbsoluteSharePointUrl(
      hyperlinkTarget.serverRelativeUrl || sourceFilePath,
    );
    const mappingList = sp.web.lists.getByTitle(listNames.loanFileMapping);
    const failedLoans: ILoanRecord[] = [];
    const duplicateLoans: ILoanRecord[] = [];
    const missingFolderLoans: ILoanRecord[] = [];
    let createdCount = 0;
    let restoredCount = 0;

    try {
      setHyperlinkLoading(true);

      for (const destinationLoan of selectedHyperlinkLoans) {
        if (!destinationLoan.id || !destinationLoan.fileRef) {
          failedLoans.push(destinationLoan);
          await logActivity({
            title: hyperlinkTarget.fileName,
            action: "Hyperlink",
            fromPath: sourceFilePath,
            toPath: "",
            loanNumber: destinationLoan.fileName || "",
            result: "Failed",
          });
          continue;
        }

        try {
          const destinationLoanPath = normalizeSharePointPath(
            destinationLoan.fileRef,
          );
          const destinationFolderPath = relativeFolderPath
            ? `${destinationLoanPath}/${relativeFolderPath}`
            : destinationLoanPath;
          let destinationFolderExists = false;

          try {
            const destinationFolder: any = await sp.web
              .getFolderByServerRelativePath(destinationFolderPath)
              .select("Exists")();
            destinationFolderExists = destinationFolder?.Exists !== false;
          } catch (folderError) {
            console.warn(
              `Destination folder not found for loan ${destinationLoan.fileName}:`,
              folderError,
            );
          }

          if (!destinationFolderExists) {
            missingFolderLoans.push(destinationLoan);
            await logActivity({
              title: hyperlinkTarget.fileName,
              action: "Hyperlink",
              fromPath: sourceFilePath,
              toPath: destinationFolderPath,
              loanNumber: destinationLoan.fileName,
              result: "Failed",
            });
            continue;
          }

          // Same as file upload: block if this file name already exists in the folder.
          const destinationFilePath = `${destinationFolderPath}/${hyperlinkTarget.fileName}`;
          let physicalFileExists = false;
          try {
            const existingFile: any = await sp.web
              .getFileByServerRelativePath(destinationFilePath)
              .select("Exists", "Name")();
            physicalFileExists =
              existingFile?.Exists !== false &&
              !!String(existingFile?.Name || "").trim();
          } catch {
            physicalFileExists = false;
          }

          const existingMappings: Pick<
            ILoanFileMapping,
            "Id" | "IsActive" | "SourceFileItemId"
          >[] = await mappingList.items
            .select("Id", "IsActive", "SourceFileItemId")
            .filter(
              `DestinationLoanItemId eq ${destinationLoan.id} and ` +
                `RelativeFolderPath eq '${escapeODataString(relativeFolderPath)}' and ` +
                `(SourceFileItemId eq ${hyperlinkTarget.id} or ` +
                `Title eq '${escapeODataString(hyperlinkTarget.fileName)}')`,
            )
            .top(5000)();
          const activeDuplicate = existingMappings.find(
            (mapping) => mapping.IsActive,
          );
          const existingSourceMapping = existingMappings.find(
            (mapping) =>
              Number(mapping.SourceFileItemId) === Number(hyperlinkTarget.id),
          );

          if (physicalFileExists || activeDuplicate) {
            duplicateLoans.push(destinationLoan);
            await logActivity({
              title: hyperlinkTarget.fileName,
              action: "Hyperlink",
              fromPath: sourceFilePath,
              toPath: destinationFolderPath,
              loanNumber: destinationLoan.fileName,
              result: "Failed",
            });
          } else if (existingSourceMapping) {
            await mappingList.items.getById(existingSourceMapping.Id).update({
              Title: hyperlinkTarget.fileName,
              SourceFileUrl: {
                Url: sourceFileUrl,
                Description: hyperlinkTarget.fileName,
              },
              IsActive: true,
            });
            restoredCount += 1;
            await logActivity({
              title: hyperlinkTarget.fileName,
              action: "Hyperlink",
              fromPath: sourceFilePath,
              toPath: destinationFolderPath,
              loanNumber: destinationLoan.fileName,
              result: "Success",
            });
          } else {
            await mappingList.items.add({
              Title: hyperlinkTarget.fileName,
              SourceFileItemId: hyperlinkTarget.id,
              SourceFileUrl: {
                Url: sourceFileUrl,
                Description: hyperlinkTarget.fileName,
              },
              DestinationLoanItemId: destinationLoan.id,
              RelativeFolderPath: relativeFolderPath,
              IsActive: true,
            });
            createdCount += 1;
            await logActivity({
              title: hyperlinkTarget.fileName,
              action: "Hyperlink",
              fromPath: sourceFilePath,
              toPath: destinationFolderPath,
              loanNumber: destinationLoan.fileName,
              result: "Success",
            });
          }

          folderCacheRef.current.delete(destinationFolderPath);
        } catch (error) {
          console.error(
            `Create file hyperlink failed for loan ${destinationLoan.fileName}:`,
            error,
          );
          failedLoans.push(destinationLoan);
          await logActivity({
            title: hyperlinkTarget.fileName,
            action: "Hyperlink",
            fromPath: sourceFilePath,
            toPath: normalizeSharePointPath(destinationLoan.fileRef),
            loanNumber: destinationLoan.fileName,
            result: "Failed",
          });
        }
      }

      const updatedCount = createdCount + restoredCount;
      if (updatedCount) {
        toastFunc(
          "success",
          "Hyperlinks saved",
          `${hyperlinkTarget.fileName} was linked to ${updatedCount} ${
            updatedCount === 1 ? "loan" : "loans"
          }.`,
        );
      }

      if (duplicateLoans.length) {
        const duplicateLoanNumbers = duplicateLoans
          .map((loan) => loan.fileName)
          .join(", ");
        toastFunc(
          "error",
          "File already uploaded",
          duplicateLoans.length === 1
            ? `"${hyperlinkTarget.fileName}" is already uploaded to this folder in loan ${duplicateLoans[0].fileName}.`
            : `"${hyperlinkTarget.fileName}" is already uploaded to this folder in loans: ${duplicateLoanNumbers}.`,
        );
      }

      if (missingFolderLoans.length) {
        const missingLoanNumbers = missingFolderLoans
          .map((loan) => loan.fileName)
          .join(", ");
        toastFunc(
          "error",
          "Destination folder not found",
          `The "${relativeFolderPath || "loan root"}" folder does not exist in ${
            missingFolderLoans.length === 1 ? "loan" : "loans"
          }: ${missingLoanNumbers}.`,
        );
      }

      if (failedLoans.length) {
        setSelectedHyperlinkLoans([
          ...missingFolderLoans,
          ...duplicateLoans,
          ...failedLoans,
        ]);
        toastFunc(
          "error",
          "Some hyperlinks failed",
          `${failedLoans.length} ${
            failedLoans.length === 1 ? "loan could" : "loans could"
          } not be updated. Please try again.`,
        );
      } else if (missingFolderLoans.length || duplicateLoans.length) {
        setSelectedHyperlinkLoans([
          ...missingFolderLoans,
          ...duplicateLoans,
        ]);
      } else {
        closeHyperlinkDialog(true);
      }
    } catch (error) {
      console.error("Create file hyperlink failed:", error);
      await logActivity({
        title: hyperlinkTarget.fileName,
        action: "Hyperlink",
        fromPath: sourceFilePath,
        toPath: "",
        result: "Failed",
      });
      toastFunc(
        "error",
        "Hyperlink failed",
        "The file could not be linked. Please try again.",
      );
    } finally {
      setHyperlinkLoading(false);
    }
  };

  const findFolder = (nodes: TreeNode[], key: string): TreeNode | null => {
    const normalizedKey = normalizeSharePointPath(key);

    for (const node of nodes) {
      if (normalizeSharePointPath(String(node.key)) === normalizedKey) {
        return node;
      }

      if (node.children) {
        const result = findFolder(node.children, key);

        if (result) {
          return result;
        }
      }
    }

    return null;
  };

  const toFolderDestination = (folder: {
    Name: string;
    ServerRelativeUrl: string;
  }): IFolderDestination => ({
    Name: folder.Name,
    ServerRelativeUrl: normalizeSharePointPath(folder.ServerRelativeUrl),
  });

  const loadDestinationFolderTree = async (
    libraryRoot: IFolderDestination,
    target: ILoanRecord | null,
  ): Promise<TreeNode[]> => {
    const rootPath = normalizeSharePointPath(libraryRoot.ServerRelativeUrl);
    const folderRows: any[] = [];
    let pagingToken: string | undefined;
    const list = sp.web.lists.getByTitle(listNames.loan);

    do {
      const response: any = await list.renderListDataAsStream({
        ViewXml: `
          <View Scope="RecursiveAll">
            <Query>
              <Where>
                <Eq>
                  <FieldRef Name="FSObjType" />
                  <Value Type="Integer">1</Value>
                </Eq>
              </Where>
              <OrderBy>
                <FieldRef Name="FileRef" Ascending="TRUE" />
              </OrderBy>
            </Query>
            <ViewFields>
              <FieldRef Name="FileRef" />
              <FieldRef Name="FileLeafRef" />
            </ViewFields>
            <RowLimit Paged="TRUE">2000</RowLimit>
          </View>
        `,
        FolderServerRelativeUrl: rootPath,
        Paging: pagingToken,
      });

      folderRows.push(...(response.Row || []));
      pagingToken = response.NextHref
        ? response.NextHref.split("?")[1]
        : undefined;
    } while (pagingToken);

    const rootNode: TreeNode = {
      key: rootPath,
      label: libraryRoot.Name || listNames.loan,
      data: libraryRoot,
      children: [],
      leaf: true,
    };
    const nodeByPath = new Map<string, TreeNode>([
      [rootPath.toLowerCase(), rootNode],
    ]);
    const isFolderTarget = target?.folderType === 1 && !!target.fileRef;

    const destinations = folderRows
      .map((row: any) =>
        toFolderDestination({
          Name: String(row.FileLeafRef || ""),
          ServerRelativeUrl: String(row.FileRef || ""),
        }),
      )
      .filter(
        (folder) =>
          !!folder.Name &&
          !!folder.ServerRelativeUrl &&
          folder.Name !== "Forms" &&
          !folder.Name.startsWith("_"),
      )
      .filter(
        (folder) =>
          !isFolderTarget ||
          !isFolderDescendantOrSelf(
            target!.fileRef,
            folder.ServerRelativeUrl,
          ),
      )
      .sort(
        (left, right) =>
          left.ServerRelativeUrl.split("/").length -
            right.ServerRelativeUrl.split("/").length ||
          left.Name.localeCompare(right.Name, undefined, {
            sensitivity: "base",
          }),
      );

    for (const destination of destinations) {
      const destinationPath = normalizeSharePointPath(
        destination.ServerRelativeUrl,
      );
      if (
        !destinationPath ||
        pathsEqual(destinationPath, rootPath) ||
        nodeByPath.has(destinationPath.toLowerCase())
      ) {
        continue;
      }

      const node: TreeNode = {
        key: destinationPath,
        label: destination.Name,
        data: destination,
        children: [],
        leaf: true,
      };
      const parentPath = getParentFolderPath(destinationPath).toLowerCase();
      const parentNode = nodeByPath.get(parentPath);
      if (!parentNode) continue;

      parentNode.children = parentNode.children || [];
      parentNode.children.push(node);
      parentNode.leaf = false;
      nodeByPath.set(destinationPath.toLowerCase(), node);
    }

    for (const node of Array.from(nodeByPath.values())) {
      if (node.children?.length) {
        node.children.sort((left, right) =>
          String(left.label || "").localeCompare(
            String(right.label || ""),
            undefined,
            { sensitivity: "base" },
          ),
        );
        node.leaf = false;
      } else {
        node.children = undefined;
        node.leaf = true;
      }
    }

    return [rootNode];
  };

  const getDestinationFromKey = (
    tree: TreeNode[],
    key: string,
    libraryRoot: IFolderDestination,
  ): IFolderDestination | null => {
    const node = findFolder(tree, key);

    if (node?.data) {
      return node.data as IFolderDestination;
    }

    if (pathsEqual(key, libraryRoot.ServerRelativeUrl)) {
      return libraryRoot;
    }

    return null;
  };

  // Library root = level 0; levels 1 and 2 are navigation-only.
  const getCopyMoveFolderLevel = (folderPath: string): number => {
    const libraryRootPath = normalizeSharePointPath(
      copyMoveLibraryRootRef.current?.ServerRelativeUrl || "",
    );
    const normalizedFolderPath = normalizeSharePointPath(folderPath);

    if (
      !libraryRootPath ||
      !isPathUnderLibrary(libraryRootPath, normalizedFolderPath)
    ) {
      return -1;
    }

    return normalizedFolderPath
      .substring(libraryRootPath.length)
      .split("/")
      .filter(Boolean).length;
  };

  const isSelectableCopyMoveDestination = (folderPath: string): boolean =>
    getCopyMoveFolderLevel(folderPath) >= 3;

  const isInvalidMoveDestination = (
    target: ILoanRecord,
    destinationPath: string,
  ): string | null => {
    const sourcePath = normalizeSharePointPath(target.fileRef);
    const destination = normalizeSharePointPath(destinationPath);
    const parentPath = getParentFolderPath(sourcePath);
    const isFile = target.folderType !== 1;

    if (!isSelectableCopyMoveDestination(destination)) {
      return "First-level and second-level folders cannot be selected.";
    }

    if (destination === parentPath) {
      return "The source and destination locations are the same.";
    }

    if (
      !isFile &&
      isFolderDescendantOrSelf(sourcePath, destination)
    ) {
      return "You can't move a folder into itself or one of its subfolders.";
    }

    if (!isPathUnderLibrary(copyMoveLibraryRootRef.current?.ServerRelativeUrl || "", destination)) {
      return "Please select a destination inside this document library.";
    }

    return null;
  };

  const isInvalidCopyDestination = (
    target: ILoanRecord,
    destinationPath: string,
  ): string | null => {
    const sourcePath = normalizeSharePointPath(target.fileRef);
    const destination = normalizeSharePointPath(destinationPath);
    const isFile = target.folderType !== 1;
    const destinationItemPath = buildDestinationItemPath(
      destination,
      target.fileName,
    );

    if (!isSelectableCopyMoveDestination(destination)) {
      return "First-level and second-level folders cannot be selected.";
    }

    if (pathsEqual(sourcePath, destinationItemPath)) {
      return "The source and destination locations are the same.";
    }

    if (
      !isFile &&
      pathsEqual(sourcePath, destination)
    ) {
      return "You can't copy a folder into itself.";
    }

    if (!isFile && isFolderDescendantOrSelf(sourcePath, destination)) {
      return "You can't copy a folder into itself or one of its subfolders.";
    }

    if (!isPathUnderLibrary(copyMoveLibraryRootRef.current?.ServerRelativeUrl || "", destination)) {
      return "Please select a destination inside this document library.";
    }

    return null;
  };

  const copyMoveNodeTemplate = (
    node: TreeNode,
    options: { expanded: boolean },
  ): React.ReactElement => {
    const isLibraryRoot = pathsEqual(
      String(node.key),
      copyMoveLibraryRootRef.current?.ServerRelativeUrl || "",
    );
    const nodePath = normalizeSharePointPath(String(node.key));
    const isSelectable = isSelectableCopyMoveDestination(nodePath);
    const isExpandable = !!node.children?.length;
    const isSelected =
      isSelectable &&
      !!selectedDestination &&
      pathsEqual(selectedDestination.ServerRelativeUrl, String(node.key));

    return (
      <span
        className={`${styles.copyMoveNode} ${isSelected ? styles.copyMoveNodeSelected : ""}`}
        data-disabled={isSelectable ? "false" : "true"}
        role="button"
        tabIndex={isSelectable || isExpandable ? 0 : -1}
        aria-disabled={!isSelectable}
        aria-label={
          isSelectable
            ? `Select ${String(node.label || "folder")} as destination`
            : `${isLibraryRoot ? "Library" : "Folder"} ${String(node.label || "")}`
        }
        onClick={(event) => {
          event.preventDefault();
          event.stopPropagation();

          if (isExpandable) {
            setCopyMoveExpandedKeys((previousKeys) => ({
              ...previousKeys,
              [nodePath]: true,
            }));
          }

          if (isSelectable) {
            onCopyMoveDestinationSelect(nodePath);
          }
        }}
        onKeyDown={(event) => {
          if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            event.stopPropagation();
            if (isExpandable) {
              setCopyMoveExpandedKeys((previousKeys) => ({
                ...previousKeys,
                [nodePath]: true,
              }));
            }
            if (isSelectable) {
              onCopyMoveDestinationSelect(nodePath);
            }
          }
        }}
      >
        <i
          className={`pi ${options.expanded ? "pi-folder-open" : "pi-folder"} ${styles.copyMoveFolderIcon}`}
        />
        <span className={styles.copyMoveNodeLabel} title={node.label}>
          {node.label}
        </span>
      </span>
    );
  };

  const openCopyMoveDialog = async (
    rowData: ILoanRecord,
    operation: "copy" | "move",
  ): Promise<void> => {
    copyMoveTargetRef.current = rowData;
    copyMoveOperationRef.current = operation;
    setCopyMoveOperation(operation);
    setCopyMoveTarget(rowData);
    setSelectedDestination(null);
    setFolderTree([]);
    setCopyMoveExpandedKeys({});
    setShowCopyMoveDialog(true);
    setCopyMoveTreeLoading(true);

    try {
      const list = sp.web.lists.getByTitle(listNames.loan);
      const [listInfo, rootFolder] = await Promise.all([
        list.select("Title")(),
        list.rootFolder.select("Name", "ServerRelativeUrl")(),
      ]);
      const libraryRoot: IFolderDestination = {
        Name: String(listInfo.Title || listNames.loan),
        ServerRelativeUrl: normalizeSharePointPath(
          rootFolder.ServerRelativeUrl,
        ),
      };
      copyMoveLibraryRootRef.current = libraryRoot;

      const tree = await loadDestinationFolderTree(libraryRoot, rowData);

      setFolderTree(tree);
      setCopyMoveExpandedKeys({
        [libraryRoot.ServerRelativeUrl]: true,
      });
    } catch (error) {
      console.error("openCopyMoveDialog error:", error);
      toastFunc("error", "Error", "Failed to load destination folders");
    } finally {
      setCopyMoveTreeLoading(false);
    }
  };

  const onCopyMoveDestinationSelect = (value: string): void => {
    const target = copyMoveTargetRef.current;
    const operation = copyMoveOperationRef.current;
    const libraryRootPath =
      copyMoveLibraryRootRef.current?.ServerRelativeUrl || "";

    if (pathsEqual(value, libraryRootPath)) {
      setSelectedDestination(null);
      return;
    }

    if (!isSelectableCopyMoveDestination(value)) {
      setSelectedDestination(null);
      return;
    }

    if (
      target?.folderType === 1 &&
      pathsEqual(value, target.fileRef)
    ) {
      toastFunc(
        "warn",
        "Warning",
        "The selected item is the source folder. Please choose a different destination.",
      );
      return;
    }

    const destination = getDestinationFromKey(
      folderTree,
      value,
      copyMoveLibraryRootRef.current || {
        Name: listNames.loan,
        ServerRelativeUrl: "",
      },
    );

    if (!destination) {
      return;
    }

    const validationMessage =
      operation === "move"
        ? isInvalidMoveDestination(target!, destination.ServerRelativeUrl)
        : isInvalidCopyDestination(target!, destination.ServerRelativeUrl);

    if (validationMessage) {
      toastFunc("warn", "Warning", validationMessage);
      return;
    }

    setSelectedDestination(destination);
  };

  const onCopyMoveNodeClick = (event: any): void => {
    const clickTarget = event.originalEvent?.target as HTMLElement | undefined;
    const key = normalizeSharePointPath(String(event.node.key));
    const isLibraryRoot = pathsEqual(
      key,
      copyMoveLibraryRootRef.current?.ServerRelativeUrl || "",
    );

    if (clickTarget?.closest("button")) {
      return;
    }

    const hasChildren = !!event.node.children?.length;

    if (hasChildren) {
      setCopyMoveExpandedKeys((prev) => ({ ...prev, [key]: true }));
    }

    if (!isLibraryRoot && isSelectableCopyMoveDestination(key)) {
      onCopyMoveDestinationSelect(key);
    }
  };

  const closeCopyMoveDialog = (): void => {
    setShowCopyMoveDialog(false);
    setCopyMoveTarget(null);
    copyMoveTargetRef.current = null;
    setSelectedDestination(null);
    setFolderTree([]);
    setCopyMoveExpandedKeys({});
    setCopyMoveTreeLoading(false);
    copyMoveLibraryRootRef.current = null;
  };

  const invalidateFolderCaches = (
    parentPath: string,
    destinationFolderPath: string,
    sourcePath: string,
    isFile: boolean,
  ): void => {
    const pathsToClear = [
      parentPath,
      destinationFolderPath,
      sourcePath,
      currentFolder,
    ]
      .filter(Boolean)
      .map((path) => normalizeSharePointPath(path));

    for (const cachedPath of Array.from(folderCacheRef.current.keys())) {
      const normalizedCachedPath = normalizeSharePointPath(cachedPath);

      for (const path of pathsToClear) {
        if (
          normalizedCachedPath === path ||
          normalizedCachedPath.startsWith(`${path}/`)
        ) {
          folderCacheRef.current.delete(cachedPath);
          break;
        }
      }
    }

    if (!isFile) {
      const normalizedSourcePath = normalizeSharePointPath(sourcePath);

      for (const cachedPath of Array.from(folderCacheRef.current.keys())) {
        const normalizedCachedPath = normalizeSharePointPath(cachedPath);

        if (
          normalizedCachedPath === normalizedSourcePath ||
          normalizedCachedPath.startsWith(`${normalizedSourcePath}/`)
        ) {
          folderCacheRef.current.delete(cachedPath);
        }
      }
    }

    pathsToClear.forEach((path) => invalidateLoanTreeCacheForPath(path));
  };

  const removeMovedItemFromView = (sourcePath: string): void => {
    const isMovedItem = (item: ILoanRecord): boolean =>
      pathsEqual(item.fileRef, sourcePath);

    setDisplayItems((prev) => prev.filter((item) => !isMovedItem(item)));

    if (!currentFolder) {
      setMainFolders((prev) => {
        const updated = prev.filter((item) => !isMovedItem(item));
        dispatch(setLoanDetails(updated));
        return updated;
      });
      return;
    }

    const cached = folderCacheRef.current.get(currentFolder);
    if (cached) {
      folderCacheRef.current.set(
        currentFolder,
        cached.filter((item) => !isMovedItem(item)),
      );
    }
  };

  const refreshViewsAfterCopyMove = async (
    parentPath: string,
    destinationFolderPath: string,
    sourcePath: string,
    operation: "copy" | "move",
  ): Promise<void> => {
    if (!currentFolder) {
      await getLoanData(drpdown.sponsor, false);
    } else {
      folderCacheRef.current.delete(currentFolder);
      invalidateLoanTreeCacheForPath(currentFolder);

      const items = await fetchFolderContents(currentFolder);
      folderCacheRef.current.set(currentFolder, items);
      await refreshCurrentView();
    }

    if (operation === "move") {
      removeMovedItemFromView(sourcePath);
    }
  };

  const executeCopyMove = async (): Promise<void> => {
    if (!selectedDestination || !copyMoveTarget) {
      toastFunc("warn", "Warning", "Please select a destination folder");
      return;
    }

    const sourcePath = copyMoveTarget.fileRef;
    const fileName = copyMoveTarget.fileName;
    const isFile = copyMoveTarget.folderType !== 1;
    const parentPath = getParentFolderPath(sourcePath);
    const destinationFolderPath = normalizeSharePointPath(
      selectedDestination.ServerRelativeUrl,
    );
    const destinationItemPath = buildDestinationItemPath(
      destinationFolderPath,
      fileName,
    );

    const validationMessage =
      copyMoveOperation === "move"
        ? isInvalidMoveDestination(copyMoveTarget, destinationFolderPath)
        : isInvalidCopyDestination(copyMoveTarget, destinationFolderPath);

    if (validationMessage) {
      toastFunc("warn", "Warning", validationMessage);
      return;
    }

    try {
      setLoader(true);

      if (copyMoveOperation === "copy") {
        if (isFile) {
          await sp.web
            .getFileByServerRelativePath(sourcePath)
            .copyByPath(destinationItemPath, true);
        } else {
          await sp.web
            .getFolderByServerRelativePath(sourcePath)
            .copyByPath(destinationItemPath);
        }
      } else if (isFile) {
        await sp.web
          .getFileByServerRelativePath(sourcePath)
          .moveByPath(destinationItemPath, true);
      } else {
        await sp.web
          .getFolderByServerRelativePath(sourcePath)
          .moveByPath(destinationItemPath);
      }

      invalidateFolderCaches(
        parentPath,
        destinationFolderPath,
        sourcePath,
        isFile,
      );
      await refreshViewsAfterCopyMove(
        parentPath,
        destinationFolderPath,
        sourcePath,
        copyMoveOperation,
      );

      await logActivity({
        title: fileName,
        action: copyMoveOperation === "copy" ? "Copy To" : "Move To",
        fromPath: sourcePath,
        toPath: destinationItemPath,
        result: "Success",
      });

      toastFunc(
        "success",
        "Success",
        `${isFile ? "File" : "Folder"} ${copyMoveOperation === "copy" ? "copied" : "moved"} successfully`,
      );
      closeCopyMoveDialog();
    } catch (error: any) {
      console.error("executeCopyMove error:", error);
      await logActivity({
        title: fileName,
        action: copyMoveOperation === "copy" ? "Copy To" : "Move To",
        fromPath: sourcePath,
        toPath: destinationItemPath,
        result: "Failed",
      });
      const rawMessage = String(error?.message || "");
      const message =
        rawMessage.includes("already exists") ||
        rawMessage.includes("SPException") ||
        rawMessage.includes("0x80070050")
          ? "A file or folder with that name already exists in the destination."
          : `Failed to ${copyMoveOperation === "copy" ? "copy" : "move"}`;
      toastFunc("error", "Error", message);
    } finally {
      setLoader(false);
    }
  };

  // Load initial sponsor and loan data on component mount.
  useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      menu.current?.hide(event as any);
      menuRef.current?.hide(event as any);
    };
    document.addEventListener("click", handleOutsideClick);
    setLoader(true);
    getSponsorData().catch((error) => console.error(error));
    return () => document.removeEventListener("click", handleOutsideClick);
  }, []);

  // Column widths for the loan table (must sum to ~100% per view).
  const isRootDashboard = !currentFolder;
  const isTaxonomyFolder =
    !!currentFolder &&
    (currentFolder.includes("Asset Management") ||
      currentFolder.includes("CSP Legal") ||
      currentFolder.includes("Servicing"));

  // Root dashboard: Name | Sponsor | Created By | Created On | Modified On | Action
  // (metadata columns only show inside taxonomy folders)
  const colWidth = isRootDashboard
    ? {
        name: "20%",
        sponsor: "18%",
        createdBy: "18%",
        createdOn: "10%",
        modifiedOn: "10%",
        action: "8%",
      }
    : isTaxonomyFolder
      ? {
          name: "22%",
          sponsor: "14%",
          meta: "20%",
          createdBy: "16%",
          createdOn: "9%",
          modifiedOn: "9%",
          action: "7%",
        }
      : {
          name: "28%",
          sponsor: "18%",
          createdBy: "22%",
          createdOn: "10%",
          modifiedOn: "10%",
          action: "8%",
        };

  const columnSize = (width: string) => ({
    style: { width, maxWidth: width },
    headerStyle: { width, maxWidth: width },
  });
  const breadcrumbSegments = getBreadcrumbSegments();
  const hiddenBreadcrumbFolderNames = breadcrumbSegments
    .slice(1, -2)
    .map((segment) => segment.label)
    .join(" > ");
  const displayedBreadcrumbSegments =
    breadcrumbSegments.length > 3
      ? [
          breadcrumbSegments[0],
          { label: "…", path: "__breadcrumb_ellipsis__" },
          ...breadcrumbSegments.slice(-2),
        ]
      : breadcrumbSegments;
  const currentBreadcrumbPath =
    breadcrumbSegments[breadcrumbSegments.length - 1]?.path || "";
  const isInsideLoanFolder = !!getActiveLoanRoot(currentFolder);
  const selectedLoanLabel = getSelectedLoanLabel();
  const isSearchActive = !!filter.search.trim();

  return (
    <>
      {loader ? (
        <Loader />
      ) : (
        <div className={styles.container}>
          {/* ── Toolbar: breadcrumb, search, sponsor filter, refresh, new ── */}
          <div className={styles.toolbar}>
            <div className={styles.toolbarTitle}>
              <span
                className={styles.breadcrumbRoot}
                onClick={goToRoot}
                title="CSP Loan Files"
              >
                CSP Loan Files
              </span>
              {displayedBreadcrumbSegments.map((segment) => {
                const isEllipsis =
                  segment.path === "__breadcrumb_ellipsis__";
                const isCurrent = segment.path === currentBreadcrumbPath;

                return (
                <React.Fragment key={segment.path}>
                  <span className={styles.breadcrumbSeparator}>&gt;&gt;</span>
                  <span
                    className={
                      isEllipsis
                        ? styles.breadcrumbCurrent
                        : isCurrent
                        ? styles.breadcrumbCurrent
                        : styles.breadcrumbLink
                    }
                    style={isEllipsis ? { flexShrink: 0, color: "#64748b" } : undefined}
                    title={
                      isEllipsis
                        ? hiddenBreadcrumbFolderNames
                        : segment.label
                    }
                    onClick={() => {
                      if (!isCurrent && !isEllipsis) {
                        goToFolder(segment.path);
                      }
                    }}
                  >
                    {segment.label}
                  </span>
                </React.Fragment>
                );
              })}
            </div>

            <div
              className={`${styles.toolbarSearchWrap} ${
                isInsideLoanFolder ? styles.toolbarSearchLoan : ""
              }`}
            >
              <div
                className={`${styles.toolbarSearch} ${
                  isInsideLoanFolder ? styles.toolbarSearchScoped : ""
                }`}
              >
                <span className={styles.searchLeading} aria-hidden="true">
                  <i className={`pi pi-search ${styles.searchIcon}`} />
                  {isInsideLoanFolder && selectedLoanLabel && (
                    <>
                      <span
                        className={styles.searchScopeChip}
                        title={`Searching loan ${selectedLoanLabel}`}
                      >
                        <i className="pi pi-folder-open" />
                        {selectedLoanLabel}
                      </span>
                      <span className={styles.searchScopeDivider} />
                    </>
                  )}
                </span>
                <InputText
                  placeholder={
                    isInsideLoanFolder
                      ? "Search folders & files..."
                      : "Search loan folders..."
                  }
                  className={styles.searchInput}
                  value={filter.search}
                  onChange={(e) => onSearchInputChange(e.target.value)}
                  aria-label={
                    isInsideLoanFolder
                      ? `Search folders and files in loan ${selectedLoanLabel}`
                      : "Search loan folders"
                  }
                />
                <div className={styles.searchTrailing}>
                  {loanTreeSearchLoading && (
                    <i
                      className={`pi pi-spin pi-spinner ${styles.searchSpinner}`}
                      aria-hidden="true"
                    />
                  )}
                  {isSearchActive && !loanTreeSearchLoading && (
                    <span className={styles.searchResultCount} aria-live="polite">
                      {displayItems.length}
                    </span>
                  )}
                  {filter.search && (
                    <button
                      type="button"
                      className={styles.searchClearBtn}
                      onClick={() => onSearchInputChange("")}
                      aria-label="Clear search"
                    >
                      <i className="pi pi-times" />
                    </button>
                  )}
                </div>
              </div>
            </div>
            <Dropdown
              options={drpdown.sponsor}
              optionLabel="name"
              placeholder="Filter by sponsor..."
              className={styles.statusDropdown}
              value={filter.sponsor}
              onChange={(e) => onSponsorChange(e.value)}
              filter
              showClear
              panelStyle={{ width: "230px" }}
            />

            <Button
              icon="pi pi-refresh"
              className={styles.refreshBtn}
              onClick={onReset}
            />
            {canShowRequestFolder && (
              <Button
                className={styles.newBtn}
                onClick={(e) => {
                  e.stopPropagation();
                  openRequestFolderDialog();
                }}
              >
                <span className={styles.newBtnLabel}>
                  <i className="pi pi-folder-plus" /> Request
                </span>
              </Button>
            )}
            {canShowNewButton && (
              <>
                <Button
                  className={styles.newBtn}
                  onClick={(e) => {
                    e.stopPropagation();
                    menu.current?.toggle(e);
                  }}
                  aria-controls="new-menu"
                  aria-haspopup
                >
                  <span className={styles.newBtnLabel}>
                    <i className="pi pi-plus" /> New
                  </span>
                  <span className={styles.newBtnDivider} />
                  <i className={`pi pi-angle-down ${styles.newBtnCaret}`} />
                </Button>

                <Menu
                  model={newMenuItems}
                  popup
                  ref={menu}
                  id="new-menu"
                  className={styles.newMenu}
                />
              </>
            )}
          </div>
          {/* ── DataTable — always reflects the CURRENT folder level only ── */}
          <div className={styles.tableContainer}>
              <DataTable
                value={displayItems}
                paginator
                rows={10}
                rowsPerPageOptions={[10, 25, 50, 100]}
                emptyMessage={
                  currentFolder
                    ? "This folder is empty."
                    : "No loan folders found."
                }
                tableStyle={{ width: "100%", tableLayout: "fixed" }}
                paginatorTemplate="CurrentPageReport RowsPerPageDropdown FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink"
                currentPageReportTemplate="Showing {first} to {last} of {totalRecords} records"
                className={styles.table}
                rowClassName={(row: ILoanRecord) =>
                  isHighlightedOriginalFile(row) ? styles.originalFileRow : ""
                }
              >
              <Column
                field="fileName"
                header={!currentFolder ? "Loan":"Subfolder"}
                body={nameBodyTemplate}
                sortable
                {...columnSize(colWidth.name)}
              />
              <Column
                field="sponsor.sponsorTitle"
                header="Sponsor"
                sortable
                {...columnSize(colWidth.sponsor)}
                body={(row: ILoanRecord) => (
                  <span
                    className={styles.truncateCell}
                    title={row.sponsor?.sponsorTitle || ""}
                  >
                    {row.sponsor?.sponsorTitle || "—"}
                  </span>
                )}
              />
              {currentFolder.includes("Asset Management") && (
                <Column
                  field="assetmanagement"
                  header="Asset Mgmt"
                  sortable
                  {...columnSize(colWidth.meta || "14%")}
                  body={(row) => metadataPillCell(row.assetmanagement)}
                />
              )}
              {currentFolder.includes("CSP Legal") && (
                <Column
                  field="legal"
                  header="Legal"
                  sortable
                  {...columnSize(colWidth.meta || "14%")}
                  body={(row) => metadataPillCell(row.legal)}
                />
              )}
              {currentFolder.includes("Servicing") && (
                <Column
                  field="servicing"
                  header="Servicing"
                  sortable
                  {...columnSize(colWidth.meta || "14%")}
                  body={(row) => metadataPillCell(row.servicing)}
                />
              )}
              <Column
                header="Created By"
                body={createdByBodyTemplate}
                sortable
                {...columnSize(colWidth.createdBy)}
              />
              <Column
                field="createddate"
                header="Created On"
                body={createdOnBodyTemplate}
                sortable
                {...columnSize(colWidth.createdOn)}
              />
              <Column
                field="modifieddate"
                header="Modified On"
                body={modifiedOnBodyTemplate}
                sortable
                {...columnSize(colWidth.modifiedOn)}
              />
              <Column
                header="Action"
                {...columnSize(colWidth.action)}
                body={actionTemplate}
              />
            </DataTable>
          </div>
          {/* Rename Dialog */}
          <Dialog
            visible={showRename}
            className={styles.renameDialog}
            style={{ width: "420px" }}
            onHide={() => setShowRename(false)}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            modal
          >
            <div className={styles.renameDialogHeader}>
              <h3 className={styles.renameDialogTitle}>Rename</h3>
            </div>

            <div className={styles.renameDialogBody}>
              <label className={styles.renameFieldLabel}>
                {selectedFolder?.folderType === 1 ? "Folder name" : "File name"}
              </label>
              <div className={styles.renameInputRow}>
                <InputText
                  className={`singlelineText ${styles.renameInput}`}
                  placeholder={
                    selectedFolder?.folderType === 1
                      ? "Enter folder name"
                      : "Enter file name"
                  }
                  value={folderName}
                  onChange={(e) => setFolderName(e.target.value)}
                />
                {selectedFolder?.folderType !== 1 && (
                  <span className={styles.renameFileExtension}>
                    {`.${selectedFolder?.fileName?.split(".").pop()}`}
                  </span>
                )}
              </div>
            </div>

            <div className={styles.renameDialogFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                onClick={() => setShowRename(false)}
              />
              <Button
                className="submitBtn"
                label="Update"
                icon="pi pi-check"
                disabled={!folderName.trim()}
                onClick={() =>
                  selectedFolder &&
                  renameFolder(selectedFolder.fileRef, folderName)
                }
              />
            </div>
          </Dialog>
          {/* Edit Tags Dialog */}
          <Dialog
            visible={showEditTagDialog}
            className={styles.editTagDialog}
            style={{ width: "560px" }}
            onHide={closeEditTagDialog}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            modal
          >
            <div className={styles.renameDialogHeader}>
              <h3 className={styles.renameDialogTitle}>Edit Tags</h3>
              {editTagTarget?.fileName && (
                <p className={styles.editTagSubtitle}>{editTagTarget.fileName}</p>
              )}
            </div>

            <div className={styles.renameDialogBody}>
              {editTagLoading ? (
                <div className={styles.editTagLoading}>Loading tags…</div>
              ) : editTagCategory ? (
                <div className={styles.editTagField}>
                  <label className={styles.uploadFieldLabel}>
                    {editTagCategory} tag
                  </label>
                  <TaxonomyTagPicker
                    key={`edit-tag-${editTagCategory}-${editTagTarget?.id ?? "none"}`}
                    category={editTagCategory}
                    tree={editTagTrees[editTagCategory] || []}
                    value={editTagTermIds[editTagCategory] || []}
                    placeholder={`Select ${editTagCategory} tags`}
                    disabled={editTagSaving}
                    onChange={(termIds) =>
                      setEditTagTermIds((prev) => ({
                        ...prev,
                        [editTagCategory]: termIds,
                      }))
                    }
                  />
                </div>
              ) : null}
            </div>

            <div className={styles.renameDialogFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                disabled={editTagSaving}
                onClick={closeEditTagDialog}
              />
              <Button
                className="submitBtn"
                label={editTagSaving ? "Updating…" : "Update"}
                icon={editTagSaving ? "pi pi-spin pi-spinner" : "pi pi-check"}
                disabled={editTagLoading || editTagSaving}
                onClick={() => void saveEditTags()}
              />
            </div>
          </Dialog>
          {/* ── Share dialog (embedded SharePoint UI) ── */}
          <Dialog
            visible={showShareDialog}
            className={styles.shareDialog}
            style={{ width: `${SHARE_DIALOG_WIDTH}px` }}
            onHide={closeShareDialog}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            resizable={false}
            modal
            dismissableMask
          >
            {shareDialogUrl && (
              <iframe
                ref={shareIframeRef}
                title={`Share ${shareItemName}`}
                src={shareDialogUrl}
                className={styles.shareDialogFrame}
                style={{
                  width: SHARE_DIALOG_WIDTH,
                  height: SHARE_DIALOG_HEIGHT,
                }}
                onLoad={() =>
                  scheduleShareDialogBranding(shareIframeRef.current)
                }
              />
            )}
          </Dialog>
          {/* copy link dialog */}
          <Dialog
            visible={showLinkDialog}
            onHide={() => setShowLinkDialog(false)}
            showHeader={false}
            draggable={false}
            resizable={false}
            modal
            style={{ width: "420px" }}
            className={styles.linkCopiedDialog}
          >
            <div className={styles.linkCopiedContent}>
              <div className={styles.linkCopiedIcon}>
                <i className="pi pi-check" />
              </div>

              <div className={styles.linkCopiedText}>
                <h3 className={styles.linkCopiedTitle}>Link copied</h3>
                <p className={styles.linkCopiedSubtitle}>
                  The sharing link is now on your clipboard.
                </p>
              </div>

              <button
                type="button"
                className={styles.linkCopiedClose}
                aria-label="Close"
                onClick={() => setShowLinkDialog(false)}
              >
                <i className="pi pi-times" />
              </button>
            </div>
          </Dialog>
          {/* ── Version History Dialog ───────────────────────────── */}
          <Dialog
            visible={showVersionHistory}
            className={
              versionActionDialog.visible
                ? `${styles.versionDialog} ${styles.deleteDialog}`
                : `${styles.versionDialog} ${styles.versionDetailsDialog}`
            }
            style={{
              width: versionActionDialog.visible ? "440px" : "640px",
              maxWidth: "94vw",
            }}
            onHide={() =>
              versionActionDialog.visible
                ? closeVersionActionDialog()
                : versionDetails
                ? setVersionDetails(null)
                : closeVersionHistoryDialog()
            }
            showHeader={false}
            draggable={false}
            modal
            closable={!versionHistoryLoading && !versionActionLoading}
          >
            {versionActionDialog.visible ? (
              <div className={styles.deleteDialogInner}>
                <div
                  className={
                    versionActionDialog.action === "restore"
                      ? styles.versionRestoreDialogIcon
                      : styles.deleteDialogIcon
                  }
                >
                  <i
                    className={
                      versionActionDialog.action === "restore"
                        ? "pi pi-replay"
                        : "pi pi-trash"
                    }
                  />
                </div>

                <h3 className={styles.deleteDialogTitle}>
                  {versionActionDialog.action === "restore"
                    ? "Restore Version"
                    : "Delete Version"}
                </h3>

                <p className={styles.deleteDialogMessage}>
                  {versionActionDialog.action === "restore"
                    ? "Are you sure you want to restore this version?"
                    : "Are you sure you want to delete this version?"}
                </p>

                <div className={styles.deleteDialogFileName}>
                  Version {versionActionDialog.version?.versionLabel || "—"}
                </div>

                <p className={styles.deleteDialogHint}>
                  {versionActionDialog.action === "restore"
                    ? "This will create a new current version from the selected file version."
                    : "This version will be permanently removed from history."}
                </p>

                <div className={styles.deleteDialogFooter}>
                  <Button
                    className="cancelBtn"
                    label="Cancel"
                    icon="pi pi-times"
                    iconPos="left"
                    disabled={versionActionLoading}
                    onClick={closeVersionActionDialog}
                  />
                  <Button
                    label={
                      versionActionLoading
                        ? versionActionDialog.action === "restore"
                          ? "Restoring…"
                          : "Deleting…"
                        : versionActionDialog.action === "restore"
                          ? "Restore"
                          : "Delete"
                    }
                    icon={
                      versionActionLoading
                        ? "pi pi-spin pi-spinner"
                        : versionActionDialog.action === "restore"
                          ? "pi pi-replay"
                          : "pi pi-trash"
                    }
                    iconPos="left"
                    className={
                      versionActionDialog.action === "restore"
                        ? styles.versionRestoreConfirm
                        : styles.deleteDialogConfirm
                    }
                    disabled={versionActionLoading}
                    onClick={() => void confirmVersionAction()}
                  />
                </div>
              </div>
            ) : versionDetails ? (
              <>
                <div className={styles.versionDetailsHeader}>
                  <div className={styles.versionDetailsHeaderContent}>
                    <span className={styles.versionDetailsHeaderIcon}>
                      <i className="pi pi-history" />
                    </span>
                    <div>
                      <h3 className={styles.versionDialogTitle}>
                      Version Details
                      </h3>
                      <p className={styles.versionDialogSubtitle}>
                        Review metadata and activity for this saved version
                      </p>
                    </div>
                  </div>
                  <button
                    type="button"
                    className={styles.versionDetailsClose}
                    aria-label="Back to version history"
                    onClick={() => setVersionDetails(null)}
                  >
                    <i className="pi pi-arrow-left" />
                  </button>
                </div>

                <div className={styles.versionDetailsBody}>
                  <div className={styles.versionDetailsSummary}>
                    <span className={styles.versionDetailsFileIcon}>
                      <i className="pi pi-file" />
                    </span>
                    <div className={styles.versionDetailsFileSummary}>
                      <button
                        type="button"
                        className={styles.versionDetailsFileName}
                        title={
                          versionDetails.name ||
                          versionHistoryTarget?.fileName ||
                          ""
                        }
                        onClick={() =>
                          openFileVersionContent(versionDetails)
                        }
                      >
                        {versionDetails.name ||
                          versionHistoryTarget?.fileName ||
                          "Unnamed document"}
                        <i className="pi pi-external-link" />
                      </button>
                      <div className={styles.versionDetailsBadges}>
                        <span className={styles.versionDetailsVersionBadge}>
                          Version {versionDetails.versionLabel || "—"}
                        </span>
                        {versionDetails.isCurrent && (
                          <span className={styles.versionCurrentBadge}>
                            Current
                          </span>
                        )}
                        {versionDetails.size &&
                          versionDetails.size !== "-" && (
                            <span className={styles.versionDetailsSize}>
                              {versionDetails.size}
                            </span>
                          )}
                      </div>
                    </div>
                  </div>

                  <div className={styles.versionDetailsPanel}>
                    <div className={styles.versionDetailsPanelTitle}>
                      <i className="pi pi-info-circle" />
                      <span>File information</span>
                    </div>
                    <div className={styles.versionDetailsGrid}>
                      <div className={styles.versionDetailsField}>
                        <span className={styles.versionMetaLabel}>Sponsor</span>
                        {versionCellTemplate(versionDetails.sponsor)}
                      </div>
                      <div className={styles.versionDetailsField}>
                        <span className={styles.versionMetaLabel}>Size</span>
                        {versionCellTemplate(versionDetails.size)}
                      </div>
                    </div>
                  </div>

                  <div className={styles.versionDetailsPanel}>
                    <div className={styles.versionDetailsPanelTitle}>
                      <i className="pi pi-clock" />
                      <span>Timeline</span>
                    </div>
                    <div className={styles.versionDetailsGrid}>
                      <div className={styles.versionDetailsField}>
                        <span className={styles.versionMetaLabel}>Created</span>
                        {versionCellTemplate(versionDetails.createddate)}
                      </div>
                      <div className={styles.versionDetailsField}>
                        <span className={styles.versionMetaLabel}>
                          Created by
                        </span>
                        {versionCellTemplate(versionDetails.createdBy)}
                      </div>
                      <div className={styles.versionDetailsField}>
                        <span className={styles.versionMetaLabel}>Modified</span>
                        {versionCellTemplate(versionDetails.modified)}
                      </div>
                      <div className={styles.versionDetailsField}>
                        <span className={styles.versionMetaLabel}>
                          Modified by
                        </span>
                        {versionCellTemplate(versionDetails.modifiedBy)}
                      </div>
                    </div>
                  </div>

                  {getVersionTagValues(versionDetails).length > 0 && (
                    <div className={styles.versionDetailsPanel}>
                      <div className={styles.versionDetailsPanelTitle}>
                        <i className="pi pi-tags" />
                        <span>Managed metadata</span>
                      </div>
                      <div className={styles.metadataTagPills}>
                        {getVersionTagValues(versionDetails).map((tag) => (
                          <span
                            key={`${versionDetails.versionLabel}-${tag}`}
                            className={styles.metadataTagPill}
                          >
                            {tag}
                          </span>
                        ))}
                      </div>
                    </div>
                  )}

                  {versionDetails.comments &&
                    versionDetails.comments !== "-" && (
                      <div className={styles.versionDetailsPanel}>
                        <div className={styles.versionDetailsPanelTitle}>
                          <i className="pi pi-comment" />
                          <span>Comments</span>
                        </div>
                        <p className={styles.versionDetailsComment}>
                          {versionDetails.comments}
                        </p>
                      </div>
                    )}
                </div>

                <div className={styles.versionDetailsFooter}>
                  <button
                    type="button"
                    className={styles.versionDetailsBackButton}
                    onClick={() => setVersionDetails(null)}
                  >
                    <i className="pi pi-arrow-left" />
                    <span>Back to History</span>
                  </button>
                  <div className={styles.versionDetailsActions}>
                    {!versionDetails.isCurrent && (
                      <>
                      <Button
                        label="Revert"
                        icon="pi pi-replay"
                        className={styles.versionDetailsRevertButton}
                        onClick={() =>
                          openVersionActionFromDetails("restore")
                        }
                      />
                      <Button
                        label="Delete"
                        icon="pi pi-trash"
                        className={styles.versionDetailsDeleteButton}
                        onClick={() =>
                          openVersionActionFromDetails("delete")
                        }
                      />
                      </>
                    )}
                  </div>
                </div>
              </>
            ) : (
              <>
                <div className={styles.versionDialogHeader}>
                  <div>
                    <h3 className={styles.versionDialogTitle}>
                      Version History
                    </h3>
                    <p className={styles.versionDialogSubtitle}>
                      View, restore, or delete versions of{" "}
                      <strong>
                        {versionHistoryTarget?.fileName || "this document"}
                      </strong>
                    </p>
                  </div>
                  {!versionHistoryLoading && (
                    <i
                      className={`pi pi-times ${styles.versionDialogClose}`}
                      onClick={closeVersionHistoryDialog}
                    />
                  )}
                </div>

                <div className={styles.versionList}>
                  {versionHistoryLoading ? (
                    <div className={styles.versionLoader}>
                      <Loader />
                    </div>
                  ) : versionHistory.length ? (
                    versionHistory.map((row) =>
                      versionHistoryRowTemplate(row),
                    )
                  ) : (
                    <div className={styles.versionEmpty}>
                      No version history found
                    </div>
                  )}
                </div>

                <div className={styles.versionDialogFooter}>
                  <Button
                    className="cancelBtn"
                    icon="pi pi-times"
                    label="Close"
                    disabled={versionHistoryLoading || versionActionLoading}
                    onClick={closeVersionHistoryDialog}
                  />
                </div>
              </>
            )}
          </Dialog>

          <Dialog
            visible={showFileUpload}
            className={styles.fileUploadDialog}
            style={{ width: "520px" }}
            onHide={closeFileUploadDialog}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            modal
            blockScroll
          >
            <div className={styles.fileUploadDialogHeader}>
              <h3 className={styles.fileUploadDialogTitle}>File Upload</h3>
            </div>

            <div className={styles.fileUploadDialogBody}>
              <div className={styles.uploadDestination} title={currentFolder}>
                <span className={styles.createFolderFieldLabel}>
                  Destination
                </span>
                <span className={styles.uploadDestinationValue}>
                  {currentFolder
                    ? currentFolder.split("/").filter(Boolean).slice(-3).join(" / ")
                    : "—"}
                </span>
              </div>

              <div
                className={`${styles.uploadArea} ${
                  uploadDragOver ? styles.uploadDragOver : ""
                } ${uploadFiles.length ? styles.uploadHasFile : ""}`}
                onDragEnter={(e) => {
                  e.preventDefault();
                  setUploadDragOver(true);
                }}
                onDragOver={(e) => {
                  e.preventDefault();
                  setUploadDragOver(true);
                }}
                onDragLeave={(e) => {
                  e.preventDefault();
                  setUploadDragOver(false);
                }}
                onDrop={onUploadDrop}
                onClick={() => {
                  if (!fileUploading) fileUploadInputRef.current?.click();
                }}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => {
                  if (e.key === "Enter" || e.key === " ") {
                    e.preventDefault();
                    fileUploadInputRef.current?.click();
                  }
                }}
              >
                <input
                  ref={fileUploadInputRef}
                  type="file"
                  multiple
                  className={styles.uploadHiddenInput}
                  onChange={onUploadFileInputChange}
                  disabled={fileUploading}
                />
                {uploadFiles.length ? (
                  <>
                    <div className={styles.uploadTitle}>
                      {uploadFiles.length} of 5 files selected
                    </div>
                    <div data-upload-file-list>
                      {uploadFiles.map((file, index) => (
                        <div
                          key={`${file.name}-${file.size}-${file.lastModified}`}
                          data-upload-file-row
                        >
                          <i className="pi pi-file" />
                          <span title={file.name}>{file.name}</span>
                          <button
                            type="button"
                            data-upload-remove-file
                            aria-label={`Remove ${file.name}`}
                            disabled={fileUploading}
                            onClick={(e) => {
                              e.stopPropagation();
                              setUploadFiles((currentFiles) =>
                                currentFiles.filter(
                                  (_, fileIndex) => fileIndex !== index,
                                ),
                              );
                            }}
                          >
                            <i className="pi pi-times" />
                          </button>
                        </div>
                      ))}
                    </div>
                    <div className={styles.uploadText}>
                      {uploadFiles.length < 5
                        ? "Click or drop more files to add"
                        : "Maximum of 5 files selected"}
                    </div>
                  </>
                ) : (
                  <>
                    <i className={`pi pi-cloud-upload ${styles.uploadIcon}`} />
                    <div className={styles.uploadTitle}>
                      Drag & drop a file here
                    </div>
                    <div className={styles.uploadText}>
                      or click to browse — maximum 5 files
                    </div>
                  </>
                )}
              </div>

              <div className={styles.uploadTagSection}>
                <label className={styles.uploadFieldLabel}>Tag category</label>
                <div className={styles.tagPills}>
                  {TAXONOMY_CATEGORIES.map((category) => (
                    <button
                      key={category}
                      type="button"
                      className={styles.tagPill}
                      data-active={
                        uploadTagCategory === category ? "true" : "false"
                      }
                      aria-pressed={uploadTagCategory === category}
                      title={
                        uploadTagCategory === category
                          ? "Click again to clear tag category"
                          : `Select ${category}`
                      }
                      disabled={fileUploading || uploadTagsLoading}
                      onClick={() => {
                        // Click active category again to clear (no tag).
                        if (uploadTagCategory === category) {
                          setUploadTagCategory(null);
                          setUploadTagTermIds([]);
                          return;
                        }
                        setUploadTagCategory(category);
                        setUploadTagTermIds([]);
                      }}
                    >
                      {category}
                    </button>
                  ))}
                </div>

                <label className={styles.uploadFieldLabel}>Tag</label>
                <TaxonomyTagPicker
                  key={`upload-${uploadTagCategory || "none"}`}
                  category={uploadTagCategory}
                  tree={
                    uploadTagCategory
                      ? uploadTagTrees[uploadTagCategory] || []
                      : []
                  }
                  value={uploadTagTermIds}
                  disabled={
                    fileUploading || uploadTagsLoading || !uploadTagCategory
                  }
                  placeholder={
                    uploadTagsLoading
                      ? "Loading tags…"
                      : "— select managed metadata tags —"
                  }
                  onChange={setUploadTagTermIds}
                />
              </div>
            </div>

            <div className={styles.fileUploadDialogFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                disabled={fileUploading}
                onClick={closeFileUploadDialog}
              />
              <Button
                label={
                  fileUploading
                    ? `Uploading ${Math.min(
                        uploadProgress.completed + 1,
                        uploadProgress.total,
                      )} of ${uploadProgress.total}…`
                    : uploadFiles.length > 1
                      ? `Upload ${uploadFiles.length} Files`
                      : "Upload"
                }
                icon={
                  fileUploading ? "pi pi-spin pi-spinner" : "pi pi-upload"
                }
                className="submitBtn"
                disabled={
                  !uploadFiles.length || fileUploading || uploadTagsLoading
                }
                onClick={() => void confirmFileUpload()}
              />
            </div>
          </Dialog>

          <Dialog
            visible={fileUploading && uploadProgress.total > 0}
            className={styles.fileUploadDialog}
            style={{ width: "520px", maxWidth: "calc(100vw - 32px)" }}
            onHide={() => undefined}
            showCloseIcon={false}
            showHeader={false}
            closable={false}
            closeOnEscape={false}
            dismissableMask={false}
            draggable={false}
            modal
            blockScroll
            appendTo={document.body}
            baseZIndex={2000}
          >
              <div data-loan-upload-progress-card>
                <div data-loan-upload-progress-header>
                  <div
                    data-loan-upload-progress-icon={
                      uploadProgress.phase === "complete"
                        ? uploadProgress.failCount
                          ? "warning"
                          : "complete"
                        : "uploading"
                    }
                  >
                    <i
                      className={
                        uploadProgress.phase === "complete"
                          ? uploadProgress.failCount
                            ? "pi pi-exclamation-triangle"
                            : "pi pi-check"
                          : "pi pi-cloud-upload"
                      }
                    />
                  </div>
                  <div>
                    <h3 id="loan-upload-progress-title">
                      {uploadProgress.phase === "complete"
                        ? uploadProgress.failCount
                          ? "Upload completed with errors"
                          : "Upload complete"
                        : "Uploading files"}
                    </h3>
                    <p>
                      {uploadProgress.phase === "complete"
                        ? `${uploadProgress.successCount} of ${uploadProgress.total} files uploaded successfully`
                        : `Processing ${Math.min(
                            uploadProgress.completed + 1,
                            uploadProgress.total,
                          )} of ${uploadProgress.total} files`}
                    </p>
                  </div>
                </div>

                <div data-loan-upload-progress-track>
                  <div
                    data-loan-upload-progress-fill
                    style={{
                      width: `${Math.round(
                        (uploadProgress.completed / uploadProgress.total) * 100,
                      )}%`,
                    }}
                  />
                </div>

                <div data-loan-upload-progress-meta>
                  <span>
                    {Math.round(
                      (uploadProgress.completed / uploadProgress.total) * 100,
                    )}
                    %
                  </span>
                  <span>
                    {uploadProgress.successCount} succeeded
                    {uploadProgress.failCount
                      ? ` • ${uploadProgress.failCount} failed`
                      : ""}
                  </span>
                </div>

                {uploadProgress.phase !== "complete" &&
                  uploadProgress.currentFileName && (
                    <div data-loan-upload-current-file>
                      <span>Current file</span>
                      <strong title={uploadProgress.currentFileName}>
                        {uploadProgress.currentFileName}
                      </strong>
                    </div>
                  )}

                <p data-loan-upload-progress-hint>
                  {uploadProgress.phase === "complete"
                    ? "Finishing up..."
                    : "Please keep this page open until the upload finishes."}
                </p>
              </div>
          </Dialog>

          {/* subfolder dialog */}
          <Dialog
            visible={showCreateFolder}
            className={styles.createFolderDialog}
            style={{ width: "420px" }}
            onHide={() => {
              if (requestFolderLoading) return;
              isRequestFolderDialogRef.current = false;
              setShowCreateFolder(false);
              setIsRequestFolderDialog(false);
            }}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            modal
          >
            <div className={styles.createFolderDialogHeader}>
              <h3 className={styles.createFolderDialogTitle}>
                {isRequestFolderDialog ? "Request Folder" : "Create Sub-Folder"}
              </h3>
            </div>

            <div className={styles.createFolderDialogBody}>
              <label className={styles.createFolderFieldLabel}>
                Folder Name <span className={styles.required}>*</span>
              </label>
              <InputText
                className={`singlelineText ${styles.createFolderInput}`}
                placeholder="Enter folder name"
                value={folderName}
                onChange={(e) => setFolderName(e.target.value)}
              />
            </div>

            <div className={styles.createFolderDialogFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                disabled={requestFolderLoading}
                onClick={() => {
                  isRequestFolderDialogRef.current = false;
                  setShowCreateFolder(false);
                  setIsRequestFolderDialog(false);
                }}
              />
              <Button
                label={
                  isRequestFolderDialog
                    ? requestFolderLoading
                      ? "Requesting…"
                      : "Request"
                    : "Create"
                }
                icon={
                  requestFolderLoading
                    ? "pi pi-spin pi-spinner"
                    : "pi pi-plus"
                }
                className="submitBtn"
                disabled={!folderName.trim() || requestFolderLoading}
                onClick={() =>
                  void (isRequestFolderDialogRef.current
                    ? submitFolderRequest()
                    : createsubFolder())
                }
              />
            </div>
          </Dialog>
          {/* ── Delete confirmation dialog ── */}
          <Dialog
            visible={showDeleteDialog}
            className={styles.deleteDialog}
            style={{ width: "440px" }}
            onHide={closeDeleteDialog}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            modal
          >
            <div className={styles.deleteDialogInner}>
              <div className={styles.deleteDialogIcon}>
                <i className="pi pi-trash" />
              </div>

              <h3 className={styles.deleteDialogTitle}>
                Delete{" "}
                {deleteTarget?.folderType === 1 ||
                selectedRowRef.current?.folderType === 1
                  ? "Folder"
                  : "File"}
              </h3>

              <p className={styles.deleteDialogMessage}>
                Are you sure you want to delete this item?
              </p>

              <div
                className={styles.deleteDialogFileName}
                title={
                  deleteTarget?.fileName ||
                  selectedRowRef.current?.fileName ||
                  "this item"
                }
              >
                {deleteTarget?.fileName ||
                  selectedRowRef.current?.fileName ||
                  "this item"}
              </div>

              {(deleteTarget?.folderType === 1 ||
                selectedRowRef.current?.folderType === 1) && (
                <p className={styles.deleteDialogHint}>
                  All files and subfolders inside it will also be removed.
                </p>
              )}

              <div className={styles.deleteDialogFooter}>
                <Button
                  className="cancelBtn"
                  label="Cancel"
                  icon="pi pi-times"
                  iconPos="left"
                  disabled={deleteLoading}
                  onClick={closeDeleteDialog}
                />
                <Button
                  label={deleteLoading ? "Deleting…" : "Delete"}
                  icon={
                    deleteLoading ? "pi pi-spin pi-spinner" : "pi pi-trash"
                  }
                  iconPos="left"
                  className={styles.deleteDialogConfirm}
                  disabled={deleteLoading}
                  onClick={() => {
                    void confirmDeleteItem();
                  }}
                />
              </div>
            </div>
          </Dialog>
          {/* ── Hyperlink file to a sponsor-associated loan ── */}
          <Dialog
            visible={showHyperlinkDialog}
            className={styles.copyMoveDialog}
            style={{ width: "520px", maxWidth: "calc(100vw - 32px)" }}
            onHide={() => closeHyperlinkDialog()}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            resizable={false}
            modal
          >
            <div data-copy-file-dialog>
              <div data-copy-file-header>
                <div data-copy-file-header-icon>
                  <i className="pi pi-link" />
                </div>
                <div>
                  <h3>Hyperlink File</h3>
                  <p>Make this file available in another loan without copying it.</p>
                </div>
              </div>

              <div data-copy-file-summary>
                <div>
                  <span>File</span>
                  <strong title={hyperlinkTarget?.fileName || ""}>
                    {hyperlinkTarget?.fileName || "—"}
                  </strong>
                </div>
                <div>
                  <span>Parent Sponsor</span>
                  <strong>
                    {hyperlinkSourceLoan?.sponsor?.sponsorTitle || "—"}
                  </strong>
                </div>
              </div>

              <div data-copy-file-field>
                <label htmlFor="copy-file-loans">
                  Loan Number <span>*</span>
                </label>
                <MultiSelect
                  inputId="copy-file-loans"
                  className={styles.hyperlinkLoanSelect}
                  panelClassName={styles.hyperlinkLoanPanel}
                  value={selectedHyperlinkLoans}
                  options={hyperlinkLoanOptions}
                  optionLabel="fileName"
                  placeholder={
                    hyperlinkLoanOptions.length
                      ? "Select loan numbers"
                      : "No other loans found for this sponsor"
                  }
                  filter
                  filterPlaceholder="Search loan numbers…"
                  display="chip"
                  maxSelectedLabels={2}
                  selectedItemsLabel="{0} loans selected"
                  showSelectAll
                  showClear
                  disabled={hyperlinkLoading || !hyperlinkLoanOptions.length}
                  itemTemplate={(loan: ILoanRecord) => (
                    <div className={styles.hyperlinkLoanOption}>
                      <i className="pi pi-folder" />
                      <span>{loan.fileName}</span>
                    </div>
                  )}
                  onChange={(event) =>
                    setSelectedHyperlinkLoans(event.value || [])
                  }
                />
                <small>
                  {selectedHyperlinkLoans.length
                    ? `${selectedHyperlinkLoans.length} of ${hyperlinkLoanOptions.length} loan${
                        hyperlinkLoanOptions.length === 1 ? "" : "s"
                      } selected`
                    : hyperlinkLoanOptions.length
                      ? `${hyperlinkLoanOptions.length} associated ${
                          hyperlinkLoanOptions.length === 1 ? "loan" : "loans"
                        } available`
                      : "There are no additional loan numbers associated with this parent sponsor."}
                </small>
              </div>

              <div data-copy-file-footer>
                <Button
                  className="cancelBtn"
                  icon="pi pi-times"
                  label="Cancel"
                  disabled={hyperlinkLoading}
                  onClick={() => closeHyperlinkDialog()}
                />
                <Button
                  className="submitBtn"
                  icon={
                    hyperlinkLoading
                      ? "pi pi-spin pi-spinner"
                      : "pi pi-link"
                  }
                  label={hyperlinkLoading ? "Saving…" : "Save"}
                  disabled={
                    hyperlinkLoading || !selectedHyperlinkLoans.length
                  }
                  onClick={() => void saveFileHyperlink()}
                />
              </div>
            </div>
          </Dialog>
          {/* ── Copy To / Move To Dialog ── */}
          <Dialog
            visible={showCopyMoveDialog}
            className={styles.copyMoveDialog}
            style={{ width: "550px" }}
            modal
            draggable={false}
            header={`${copyMoveOperation === "copy" ? "Copy" : "Move"} "${copyMoveTarget?.fileName}"`}
            onHide={closeCopyMoveDialog}
            showCloseIcon={false}
            showHeader={false}
            
          >
             <h3 className="modelHeader">{`${copyMoveOperation === "copy" ? "Copy" : "Move"} to`}</h3>

            <div className={styles.copyMoveSource}>
              <span className={styles.copyMoveSourceLabel}>Source</span>
              <span className={styles.copyMoveSourceValue}>
                <i
                  className={`pi ${copyMoveTarget?.folderType === 1 ? "pi-folder" : "pi-file"} ${styles.copyMoveSourceIcon}`}
                />
                <strong>{copyMoveTarget?.fileName}</strong>
              </span>
            </div>

            <p className={styles.copyMoveSectionLabel}>Destination folder</p>
            <p className={styles.copyMoveHint}>
              Expand the first-level and second-level folders, then select a
              folder below them.
            </p>

            <div className={styles.copyMoveTreePanel}>
              {copyMoveTreeLoading ? (
                <div className={styles.copyMoveStatus}>
                  <i className="pi pi-spin pi-spinner" />
                  Loading folders...
                </div>
              ) : folderTree.length === 0 ? (
                <div className={styles.copyMoveStatus}>
                  No folders available
                </div>
              ) : (
                <Tree
                  className={styles.copyMoveTree}
                  value={folderTree}
                  expandedKeys={copyMoveExpandedKeys}
                  onToggle={(e: any) => setCopyMoveExpandedKeys(e.value)}
                  metaKeySelection={false}
                  nodeTemplate={copyMoveNodeTemplate}
                  onNodeClick={onCopyMoveNodeClick}
                />
              )}
            </div>

            {selectedDestination && (
              <div className={styles.copyMoveDestination}>
                <i className="pi pi-folder" />
                <span>
                  Destination: <strong>{selectedDestination.Name}</strong>
                </span>
              </div>
            )}

            <div className={styles.copyMoveFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                onClick={closeCopyMoveDialog}
              />
              <Button
                label={copyMoveOperation === "copy" ? "Copy Here" : "Move Here"}
                icon={
                  copyMoveOperation === "copy"
                    ? "pi pi-copy"
                    : "pi pi-arrow-right-arrow-left"
                }
                className="submitBtn"
                disabled={!selectedDestination}
                onClick={executeCopyMove}
              />
            </div>
          </Dialog>
          {/* ── Row action menu (single shared instance) ── */}
          <Menu model={rowMenuItems} popup ref={menuRef} className={styles.newMenu} />
        </div>
      )}
    </>
  );
};

export default Loan;
