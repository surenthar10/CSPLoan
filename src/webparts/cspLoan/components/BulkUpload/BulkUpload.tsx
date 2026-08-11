/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { Button } from "primereact/button";
import { Dropdown } from "primereact/dropdown";
import { Tree } from "primereact/tree";
import { TreeNode } from "primereact/treenode";
import { OverlayPanel } from "primereact/overlaypanel";
import styles from "./BulkUpload.module.scss";
import {
  bulkUploadConfig,
  listNames,
  managedMetadataFields,
  toastFunc,
} from "../../assets/Config/Config";
import {
  IBulkTaxonomyCategory,
  ILoanFolder,
  IPathOption,
  ITaxonomyFieldInfo,
  ITaxonomyTag,
  ITaxonomyTagPickerProps,
  ITermStoreTerm,
  IUploadFileItem,
  IUploadProgress,
  IUploadResult,
} from "../../assets/Config/interface";
import Loader from "../Loader/Loader";

const MAX_FILES = bulkUploadConfig.maxFiles;
const UPLOAD_CONCURRENCY = bulkUploadConfig.uploadConcurrency;

const METADATA_TAG_CATEGORIES = {
  "Asset Management": [managedMetadataFields.assetManagement],
  Legal: [managedMetadataFields.legal],
  Servicing: [managedMetadataFields.servicing],
} as const;

type TaxonomyCategory = IBulkTaxonomyCategory;

const TAXONOMY_CATEGORIES = Object.keys(
  METADATA_TAG_CATEGORIES,
) as TaxonomyCategory[];

const normalizeFolderPath = (path: string): string =>
  path.replace(/\\/g, "/").replace(/\/+$/, "");

// Normalize and decode SharePoint path values.
const normalizeSharePointPath = (path: string): string => {
  if (!path) return "";
  try {
    return normalizeFolderPath(decodeURIComponent(path));
  } catch {
    return normalizeFolderPath(path);
  }
};

const areSharePointPathsEqual = (left: string, right: string): boolean =>
  normalizeSharePointPath(left).toLowerCase() ===
  normalizeSharePointPath(right).toLowerCase();

const isPathInsideFolder = (path: string, folderPath: string): boolean => {
  const normalizedPath = normalizeSharePointPath(path).toLowerCase();
  const normalizedFolder = normalizeSharePointPath(folderPath).toLowerCase();
  return (
    normalizedPath === normalizedFolder ||
    normalizedPath.startsWith(`${normalizedFolder}/`)
  );
};

const shouldSkipFolderName = (name: string): boolean =>
  name === "Forms" || name.startsWith("_");

const isValidFolderPath = (
  folderPath: string,
  loanRootPath: string,
): boolean => {
  const relative = folderPath
    .slice(normalizeSharePointPath(loanRootPath).length)
    .replace(/^\//, "");
  if (!relative) return true;
  return relative
    .split("/")
    .every((segment) => segment && !shouldSkipFolderName(segment));
};

const buildFolderPathDisplayLabel = (
  loanName: string,
  relativePath: string,
): string =>
  relativePath
    ? `${loanName}-${relativePath.replace(/\//g, "-")}`
    : loanName;

const normalizeTextForMatching = (value: string): string =>
  value.toLowerCase().replace(/[^a-z0-9]/g, "");

// Retry short-lived SharePoint failures once before showing an error.
const withSharePointRetry = async <T,>(
  operation: () => Promise<T>,
): Promise<T> => {
  try {
    return await operation();
  } catch {
    await new Promise((resolve) => setTimeout(resolve, 500));
    return operation();
  }
};

// Resolve the automatic tag category from the selected destination path.
const determineTagCategoryFromFolderPath = (
  folderPath: string,
): TaxonomyCategory | null => {
  const normalizedPath = normalizeTextForMatching(
    normalizeSharePointPath(folderPath),
  );

  if (normalizedPath.includes("assetmanagement")) return "Asset Management";
  if (normalizedPath.includes("legal")) return "Legal";
  if (normalizedPath.includes("servicing")) return "Servicing";
  return null;
};

// Build all valid folder paths for a loan from one recursive list result.
const buildFolderPathOptionsForLoan = (
  loanFolder: ILoanFolder,
  folderRows: any[],
): IPathOption[] => {
  const loanRootPath = normalizeSharePointPath(loanFolder.fileRef);
  const folderPathOptions: IPathOption[] = [
    {
      key: loanRootPath,
      label: loanFolder.name,
      loanNumber: loanFolder.name,
      folderPath: loanRootPath,
    },
  ];

  const visitedFolderPaths = new Set<string>([loanRootPath.toLowerCase()]);

  for (const folderRow of folderRows) {
    const folderPath = normalizeSharePointPath(folderRow.FileRef || "");
    const folderName = folderRow.FileLeafRef || "";
    const normalizedFolderPath = folderPath.toLowerCase();

    if (!folderPath || visitedFolderPaths.has(normalizedFolderPath)) continue;
    if (shouldSkipFolderName(folderName)) continue;
    if (!isValidFolderPath(folderPath, loanRootPath)) continue;
    if (!isPathInsideFolder(folderPath, loanRootPath)) continue;

    const relativeFolderPath = folderPath
      .slice(loanRootPath.length)
      .replace(/^\//, "");

    visitedFolderPaths.add(normalizedFolderPath);
    folderPathOptions.push({
      key: folderPath,
      label: buildFolderPathDisplayLabel(
        loanFolder.name,
        relativeFolderPath,
      ),
      loanNumber: loanFolder.name,
      folderPath,
    });
  }

  return folderPathOptions;
};

// Merge and sort folder path options across loans.
const loadFolderPathOptionsFromSharePoint = async (
  loanFolders: ILoanFolder[],
): Promise<IPathOption[]> => {
  if (!loanFolders.length) return [];

  const loanDocumentLibrary = sp.web.lists.getByTitle(listNames.loan);
  const libraryRootFolder = await loanDocumentLibrary.rootFolder();
  const folderRows: any[] = [];
  let pagingToken: string | undefined;

  do {
    const folderPageResponse: any =
      await loanDocumentLibrary.renderListDataAsStream({
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
      FolderServerRelativeUrl: libraryRootFolder.ServerRelativeUrl,
      Paging: pagingToken,
    });

    folderRows.push(...(folderPageResponse.Row || []));
    pagingToken = folderPageResponse.NextHref
      ? folderPageResponse.NextHref.split("?")[1]
      : undefined;
  } while (pagingToken);

  const folderOptionsByLoan = loanFolders.map((loanFolder) =>
    buildFolderPathOptionsForLoan(loanFolder, folderRows),
  );

  const mergedFolderPathOptions: IPathOption[] = [];
  for (const loanFolderOptions of folderOptionsByLoan) {
    mergedFolderPathOptions.push(...loanFolderOptions);
  }

  return mergedFolderPathOptions.sort((a: IPathOption, b: IPathOption) =>
    a.label.localeCompare(b.label, undefined, { sensitivity: "base" }),
  );
};

// Suggest a best-matching folder path from file name.
const suggestFolderPathFromFileName = (
  fileName: string,
  pathOptions: IPathOption[],
): string => {
  const stem = fileName.replace(/\.[^/.]+$/, "");
  const normalizedStem = normalizeTextForMatching(stem);
  const fileLoanPrefix = stem.match(/^\s*([a-z0-9]+)/i)?.[1] || "";
  const normalizedPrefix = normalizeTextForMatching(fileLoanPrefix);

  const loanNumbers = Array.from(
    new Set(pathOptions.map((option) => option.loanNumber.trim())),
  ).sort((a, b) => b.length - a.length);
  const loanNumber = loanNumbers.find((candidate) => {
    const normalizedLoan = normalizeTextForMatching(candidate);
    return (
      normalizedStem.startsWith(normalizedLoan) ||
      (!!normalizedPrefix &&
        (normalizedLoan === normalizedPrefix ||
          normalizedLoan.startsWith(normalizedPrefix) ||
          normalizedLoan.includes(normalizedPrefix)))
    );
  });
  if (!loanNumber) return "";

  const normalizedLoan = normalizeTextForMatching(loanNumber);
  const loanOptions = pathOptions.filter(
    (option) => normalizeTextForMatching(option.loanNumber) === normalizedLoan,
  );
  if (!loanOptions.length) return "";

  const rootOption = loanOptions.reduce((shortest, option) =>
    option.folderPath.length < shortest.folderPath.length ? option : shortest,
  );
  const loanPosition = normalizedStem.indexOf(normalizedLoan);
  const filePathHint =
    loanPosition >= 0
      ? normalizedStem.slice(loanPosition + normalizedLoan.length)
      : normalizedStem.slice(normalizedPrefix.length);
  let bestKey = rootOption.key;
  let bestScore = 0;

  for (const option of loanOptions) {
    const relativePath = normalizeSharePointPath(option.folderPath)
      .slice(normalizeSharePointPath(rootOption.folderPath).length)
      .replace(/^\//, "");
    const optionHint = normalizeTextForMatching(relativePath);
    if (
      optionHint &&
      filePathHint.includes(optionHint) &&
      optionHint.length > bestScore
    ) {
      bestScore = optionHint.length;
      bestKey = option.key;
    }
  }

  return bestKey;
};

// Format a readable label for a server-relative folder path.
const formatFolderPathLabel = (
  folderPath: string,
  loanFolders: ILoanFolder[],
): string => {
  if (!folderPath) return "";
  const normalized = normalizeSharePointPath(folderPath);
  const loan = loanFolders.find(
    (l) =>
      areSharePointPathsEqual(normalized, l.fileRef) ||
      normalized.startsWith(`${normalizeSharePointPath(l.fileRef)}/`),
  );

  if (!loan) {
    return normalized.split("/").filter(Boolean).slice(-2).join("-");
  }

  if (areSharePointPathsEqual(normalized, loan.fileRef)) return loan.name;

  const relative = normalized
    .slice(normalizeSharePointPath(loan.fileRef).length)
    .replace(/^\//, "");
  return `${loan.name}-${relative.replace(/\//g, "-")}`;
};

const normalizeGuidValue = (guid: string | undefined): string =>
  (guid || "").replace(/[{}]/g, "").trim().toLowerCase();

// Compare term ids without case sensitivity.
const areTermIdsEqual = (left: string, right: string): boolean =>
  left.trim().toLowerCase() === right.trim().toLowerCase();

const getNormalizedTermId = (term: ITermStoreTerm): string =>
  normalizeGuidValue(term.id || term.Id);

const getDefaultTermLabel = (term: ITermStoreTerm): string => {
  const labels = term.labels || [];
  const preferred =
    labels.find((label) => label.isDefault)?.name || labels[0]?.name || "";
  return preferred.trim();
};

// Map a term object to a PrimeReact tree node.
const mapTermStoreTermToTreeNode = (term: ITermStoreTerm): TreeNode => {
  const termId = getNormalizedTermId(term);
  const children = (term.children || []).map(mapTermStoreTermToTreeNode);
  const label = getDefaultTermLabel(term);
  return {
    key: termId,
    label,
    data: {
      termId,
      label,
    },
    children,
    leaf: children.length === 0,
    selectable: true,
  };
};

// Build a nested tree structure from flat term responses.
const buildTermTreeFromFlatTerms = (terms: ITermStoreTerm[]): TreeNode[] => {
  const termMap = new Map<string, TreeNode>();
  const roots: TreeNode[] = [];

  for (const term of terms) {
    const termId = getNormalizedTermId(term);
    if (!termId) continue;

    const label = getDefaultTermLabel(term);
    termMap.set(termId, {
      key: termId,
      label,
      data: { termId, label },
      children: [],
      leaf: true,
      selectable: true,
    });
  }

  for (const term of terms) {
    const termId = getNormalizedTermId(term);
    if (!termId) continue;

    const node = termMap.get(termId);
    if (!node) continue;

    const parentId = normalizeGuidValue(term.parent?.id);
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

// Fetch JSON safely from a term store endpoint.
const fetchTermStoreJsonResponse = async (
  url: string,
): Promise<any | null> => {
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json;odata=nometadata",
    },
    credentials: "same-origin",
  });

  if (!response.ok) return null;
  return response.json();
};

// Load child terms from legacy term store endpoint.
const loadLegacyChildTerms = async (
  siteUrl: string,
  termSetId: string,
  parentTermId?: string,
): Promise<ITermStoreTerm[]> => {
  const setId = normalizeGuidValue(termSetId);
  const parentPath = parentTermId
    ? `/terms/${normalizeGuidValue(parentTermId)}/getlegacychildren`
    : "/getlegacychildren";
  const payload = await fetchTermStoreJsonResponse(
    `${siteUrl}/_api/v2.1/termStore/termSets/${setId}${parentPath}?$select=id,labels,childrenCount`,
  );

  return payload?.value || [];
};

// Recursively build tree nodes using legacy children endpoint.
const loadLegacyTermSetTree = async (
  siteUrl: string,
  termSetId: string,
  parentTermId?: string,
): Promise<TreeNode[]> => {
  const terms = await loadLegacyChildTerms(siteUrl, termSetId, parentTermId);
  const nodes: TreeNode[] = [];

  for (const term of terms) {
    const termId = getNormalizedTermId(term);
    if (!termId) continue;

    const children =
      (term.childrenCount || 0) > 0
        ? await loadLegacyTermSetTree(siteUrl, termSetId, termId)
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

// Load term set tree with modern endpoints and fallback logic.
const loadTermSetTreeNodes = async (
  termSetId: string,
  groupId?: string,
): Promise<TreeNode[]> => {
  const { Url } = await sp.web.select("Url")();
  const setId = normalizeGuidValue(termSetId);
  let group = groupId ? normalizeGuidValue(groupId) : "";

  if (!group) {
    const setMeta = await fetchTermStoreJsonResponse(
      `${Url}/_api/v2.1/termStore/sets/${setId}`,
    );
    group = normalizeGuidValue(setMeta?.parent?.id || setMeta?.groupId);
  }

  const urls = [
    group
      ? `${Url}/_api/v2.1/termStore/groups/${group}/sets/${setId}/terms?$expand=children($levels=max)`
      : "",
    `${Url}/_api/v2.1/termStore/sets/${setId}/terms?$expand=children($levels=max)`,
    `${Url}/_api/v2.1/termStore/sets/${setId}/terms?$top=500`,
  ].filter(Boolean);

  for (const url of urls) {
    const payload = await fetchTermStoreJsonResponse(url);
    const terms: ITermStoreTerm[] = payload?.value || [];
    if (!terms.length) continue;

    const hasNestedChildren = terms.some(
      (term) => (term.children || []).length > 0,
    );
    if (hasNestedChildren) {
      return terms.map(mapTermStoreTermToTreeNode);
    }

    if (terms.some((term) => term.parent?.id)) {
      const flatTree = buildTermTreeFromFlatTerms(terms);
      if (flatTree.length) return flatTree;
    }

    return terms.map(mapTermStoreTermToTreeNode);
  }

  return loadLegacyTermSetTree(Url, setId);
};

// Resolve taxonomy field metadata for one category.
const resolveTaxonomyFieldForCategory = async (
  category: TaxonomyCategory,
): Promise<ITaxonomyFieldInfo | null> => {
  const candidates = METADATA_TAG_CATEGORIES[category];

  for (const name of candidates) {
    try {
      const field: any = await sp.web.lists
        .getByTitle(listNames.loan)
        .fields.getByInternalNameOrTitle(name)
        .select("InternalName", "TermSetId", "SspId", "TypeAsString")();

      if (field?.TermSetId) {
        return {
          category,
          internalName: field.InternalName,
          termSetId: normalizeGuidValue(field.TermSetId),
          groupId: field.SspId ? normalizeGuidValue(field.SspId) : undefined,
        };
      }
    } catch {
      // Try the next known internal name for this category.
    }
  }

  return null;
};

// Resolve taxonomy fields for all configured categories.
const resolveTaxonomyFieldsForCategories = async (): Promise<
  ITaxonomyFieldInfo[]
> => {
  const results = await Promise.all(
    TAXONOMY_CATEGORIES.map((category) =>
      resolveTaxonomyFieldForCategory(category),
    ),
  );

  return results.filter(
    (field): field is ITaxonomyFieldInfo => field !== null,
  );
};

// Load taxonomy trees and field mappings for each category.
const loadManagedMetadataTagTrees = async (): Promise<{
  trees: Record<TaxonomyCategory, TreeNode[]>;
  fields: Record<TaxonomyCategory, string>;
}> => {
  const trees = {} as Record<TaxonomyCategory, TreeNode[]>;
  const fields = {} as Record<TaxonomyCategory, string>;
  const fieldInfos = await resolveTaxonomyFieldsForCategories();

  await Promise.all(
    fieldInfos.map(async (info) => {
      try {
        trees[info.category] = await loadTermSetTreeNodes(
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

type ManagedMetadataResult = Awaited<
  ReturnType<typeof loadManagedMetadataTagTrees>
>;

let managedMetadataCache: Promise<ManagedMetadataResult> | null = null;

// Reuse taxonomy metadata when users switch away from and back to this tab.
const loadManagedMetadataTagTreesCached =
  (): Promise<ManagedMetadataResult> => {
    if (!managedMetadataCache) {
      managedMetadataCache = loadManagedMetadataTagTrees().catch((error) => {
        managedMetadataCache = null;
        throw error;
      });
    }
    return managedMetadataCache;
  };

// Find a tree node by taxonomy term id.
const findTermInTree = (nodes: TreeNode[], termId: string): TreeNode | null => {
  for (const node of nodes) {
    if (areTermIdsEqual(String(node.key), termId)) return node;
    if (node.children?.length) {
      const match = findTermInTree(node.children, termId);
      if (match) return match;
    }
  }
  return null;
};

// Resolve the display label for a selected term id.
const findTermLabelById = (nodes: TreeNode[], termId: string): string => {
  for (const node of nodes) {
    if (areTermIdsEqual(String(node.key), termId)) {
      return String(node.label || node.data?.label || "");
    }
    if (node.children?.length) {
      const childLabel = findTermLabelById(node.children, termId);
      if (childLabel) return childLabel;
    }
  }
  return "";
};

// Build upload-ready taxonomy tag payload from selected term.
const buildTaxonomyTagPayload = (
  termId: string,
  category: TaxonomyCategory,
  trees: Record<TaxonomyCategory, TreeNode[]>,
  fields: Record<TaxonomyCategory, string>,
): ITaxonomyTag | null => {
  const node = findTermInTree(trees[category] || [], termId);
  if (!node?.data || !fields[category]) return null;

  return {
    termId: String(node.key),
    label: node.data.label || String(node.label),
    fieldInternalName: fields[category],
    category,
  };
};

// Build taxonomy tag payloads from multiple selected term ids.
const buildTaxonomyTagPayloads = (
  termIds: string[],
  category: TaxonomyCategory,
  trees: Record<TaxonomyCategory, TreeNode[]>,
  fields: Record<TaxonomyCategory, string>,
): ITaxonomyTag[] =>
  termIds
    .map((termId) => {
      const tag = buildTaxonomyTagPayload(termId, category, trees, fields);
      if (tag) return tag;

      const fallbackLabel = findTermLabelById(
        trees[category] || [],
        termId,
      );
      if (!fields[category]) return null;

      return {
        termId,
        label: fallbackLabel || termId,
        fieldInternalName: fields[category],
        category,
      };
    })
    .filter((tag): tag is ITaxonomyTag => tag !== null);

// Find a taxonomy node by its displayed label.
const findTermNodeByLabel = (
  nodes: TreeNode[],
  label: string,
): TreeNode | null => {
  const normalizedLabel = normalizeTextForMatching(label);

  for (const node of nodes) {
    const nodeLabel = String(node.data?.label || node.label || "");
    if (normalizeTextForMatching(nodeLabel) === normalizedLabel) return node;

    if (node.children?.length) {
      const childMatch = findTermNodeByLabel(node.children, label);
      if (childMatch) return childMatch;
    }
  }

  return null;
};

// Load the three automatic terms from the shared managed-metadata Tags field.
const loadAutomaticFolderTags = async (): Promise<
  Partial<Record<TaxonomyCategory, ITaxonomyTag>>
> => {
  const field: any = await sp.web.lists
    .getByTitle(listNames.loan)
    .fields.getByInternalNameOrTitle(managedMetadataFields.tags)
    .select("InternalName", "TermSetId", "SspId")();

  if (!field?.InternalName || !field?.TermSetId) {
    throw new Error("The Tags managed metadata field is not configured");
  }

  const tree = await loadTermSetTreeNodes(
    normalizeGuidValue(field.TermSetId),
    field.SspId ? normalizeGuidValue(field.SspId) : undefined,
  );
  const tags: Partial<Record<TaxonomyCategory, ITaxonomyTag>> = {};

  for (const category of TAXONOMY_CATEGORIES) {
    const term = findTermNodeByLabel(tree, category);
    const termId = String(term?.key || "");
    if (!term || !termId) continue;

    tags[category] = {
      termId,
      label: String(term.data?.label || term.label || category),
      fieldInternalName: field.InternalName,
      category,
    };
  }

  return tags;
};

type AutomaticFolderTagsResult = Awaited<
  ReturnType<typeof loadAutomaticFolderTags>
>;

let automaticFolderTagsCache: Promise<AutomaticFolderTagsResult> | null = null;

const loadAutomaticFolderTagsCached =
  (): Promise<AutomaticFolderTagsResult> => {
    if (!automaticFolderTagsCache) {
      automaticFolderTagsCache = loadAutomaticFolderTags().catch((error) => {
        automaticFolderTagsCache = null;
        throw error;
      });
    }
    return automaticFolderTagsCache;
  };

// Convert a selected term into SharePoint update wire format.
const formatTaxonomyFieldWireValue = (
  label: string,
  termId: string,
): string => {
  const termGuid = termId.replace(/[{}]/g, "").trim();
  return `${label}|${termGuid}`;
};

// Extract any field update failures from validateUpdate response.
const extractValidateUpdateFailures = (result: any): any[] => {
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

// Apply one or more taxonomy values to a list item with fallback formats.
const applyTaxonomyTagsToListItem = async (
  itemId: number,
  tags: ITaxonomyTag[],
): Promise<void> => {
  if (!tags.length) return;

  const fieldInternalName = tags[0].fieldInternalName;
  if (!fieldInternalName) {
    throw new Error(`No SharePoint field mapped for ${tags[0].category}`);
  }

  const wireValues = tags.map((tag) =>
    formatTaxonomyFieldWireValue(tag.label, tag.termId),
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
          [
            {
              FieldName: fieldInternalName,
              FieldValue: fieldValue,
            },
          ],
          true,
        );

      const failures = extractValidateUpdateFailures(result);
      if (!failures.length) {
        return;
      }

      lastError =
        failures[0]?.ErrorMessage ||
        failures[0]?.Message ||
        lastError;
    } catch (error) {
      lastError = error instanceof Error ? error.message : String(error);
    }
  }

  throw new Error(lastError);
};

// Run async tasks with a fixed concurrency limit.
const runTasksWithConcurrency = async <T, R>(
  items: T[],
  concurrency: number,
  worker: (item: T) => Promise<R>,
): Promise<R[]> => {
  if (!items.length) return [];

  const results: R[] = new Array(items.length);
  let nextIndex = 0;

  const executeWorker = async (): Promise<void> => {
    while (nextIndex < items.length) {
      const currentIndex = nextIndex++;
      results[currentIndex] = await worker(items[currentIndex]);
    }
  };

  await Promise.all(
    Array.from({ length: Math.min(concurrency, items.length) }, executeWorker),
  );

  return results;
};

// Upload one file and apply managed metadata when tags are present.
const uploadFileWithMetadata = async (
  item: IUploadFileItem,
): Promise<IUploadResult> => {
  try {
    const result = await sp.web
      .getFolderByServerRelativePath(item.folderPath)
      .files.add(item.fileName, item.file, true);

    if (item.tags.length) {
      const fileItem: any = await result.file.listItemAllFields();
      const itemId = Number(fileItem.Id ?? fileItem.ID);

      if (!itemId) {
        throw new Error(`Could not resolve list item ID for ${item.fileName}`);
      }

      await applyTaxonomyTagsToListItem(itemId, item.tags);
    }

    return { fileName: item.fileName, success: true };
  } catch (error) {
    return {
      fileName: item.fileName,
      success: false,
      error: error instanceof Error ? error.message : "Upload failed",
    };
  }
};

// Render taxonomy picker with multi-select tree.
const TaxonomyTagPicker = ({
  category,
  tree,
  value,
  displayLabels: displayLabelsProp,
  placeholder = "— select tag —",
  disabled = false,
  onChange,
}: ITaxonomyTagPickerProps): React.ReactElement => {
  const overlayRef = useRef<OverlayPanel>(null);
  const [selectedTermIds, setSelectedTermIds] = useState<string[]>(value);
  const [selectedLabels, setSelectedLabels] = useState<string[]>([]);
  const [selectionKeys, setSelectionKeys] = useState<Record<string, boolean>>({});

  const syncSelectionFromValue = useCallback(
    (termIds: string[], labelsOverride?: string[]): void => {
      setSelectedTermIds(termIds);

      if (!termIds.length) {
        setSelectedLabels([]);
        setSelectionKeys({});
        return;
      }

      const keys: Record<string, boolean> = {};
      const labels: string[] = [];

      termIds.forEach((termId, index) => {
        const treeKey = String(findTermInTree(tree, termId)?.key || termId);
        keys[treeKey] = true;
        labels.push(
          labelsOverride?.[index] ||
            displayLabelsProp?.[index] ||
            findTermLabelById(tree, termId) ||
            "",
        );
      });

      setSelectedLabels(labels.filter(Boolean));
      setSelectionKeys(keys);
    },
    [displayLabelsProp, tree],
  );

  useEffect(() => {
    syncSelectionFromValue(value);
  }, [value, syncSelectionFromValue]);

  const resolvedLabel =
    selectedLabels.join(", ") ||
    value
      .map((termId) => findTermLabelById(tree, termId))
      .filter(Boolean)
      .join(", ") ||
    (displayLabelsProp || []).join(", ");
  const hasSelection = selectedTermIds.length > 0;

  const handleToggle = (termId: string, label: string): void => {
    const isSelected = selectedTermIds.some((id) =>
      areTermIdsEqual(id, termId),
    );
    const nextIds = isSelected
      ? selectedTermIds.filter((id) => !areTermIdsEqual(id, termId))
      : [...selectedTermIds, termId];

    syncSelectionFromValue(
      nextIds,
      isSelected
        ? selectedLabels.filter(
            (_, index) => !areTermIdsEqual(selectedTermIds[index], termId),
          )
        : [...selectedLabels, label],
    );
    onChange(nextIds);
  };

  const nodeTemplate = (
    node: TreeNode,
  ): React.ReactElement => {
    const termId = String(node.key);
    const label = String(node.label || node.data?.label || "");
    const isSelected = selectedTermIds.some((id) =>
      areTermIdsEqual(id, termId),
    );

    return (
      <span
        role="button"
        tabIndex={0}
        data-selected={isSelected ? "true" : "false"}
        className={`${styles.copyMoveNode} ${
          isSelected ? styles.copyMoveNodeSelected : ""
        }`}
        onClick={(e) => {
          e.stopPropagation();
          handleToggle(termId, label);
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            handleToggle(termId, label);
          }
        }}
      >
        <i className={`pi pi-tag ${styles.copyMoveTagIcon}`} />
        <span className={styles.copyMoveNodeLabel} title={label}>
          {label}
        </span>
      </span>
    );
  };

  return (
    <>
      <button
        type="button"
        className={styles.taxonomyTagPicker}
        data-has-selection={hasSelection ? "true" : "false"}
        data-category-selected={category ? "true" : "false"}
        disabled={disabled || !category}
        onClick={(e) => overlayRef.current?.toggle(e)}
      >
        <span
          className={
            hasSelection ? styles.taxonomyTagValue : styles.pathPlaceholder
          }
        >
          {!category
            ? "Select a tag category first"
            : !tree.length
              ? `No ${category} terms found`
              : resolvedLabel || placeholder}
        </span>
        <i className={`pi pi-angle-down ${styles.pathPickerIcon}`} />
      </button>

      <OverlayPanel ref={overlayRef} className={styles.taxonomyOverlay}>
        <div className={styles.taxonomyOverlayTitle}>
          {category ? `${category} tags` : "Tags"}
        </div>
        <div className={styles.taxonomyTreePanel}>
          {tree.length ? (
            <Tree
              className={styles.copyMoveTree}
              value={tree}
              nodeTemplate={nodeTemplate}
              selectionMode="multiple"
              selectionKeys={selectionKeys}
              metaKeySelection={false}
            />
          ) : (
            <div className={styles.copyMoveStatus}>No terms available</div>
          )}
        </div>
      </OverlayPanel>
    </>
  );
};

// Pick icon class based on file extension.
const getFileIconClassName = (fileName: string): string => {
  const ext = fileName.split(".").pop()?.toLowerCase();
  switch (ext) {
    case "pdf":
      return "pi pi-file-pdf";
    case "doc":
    case "docx":
      return "pi pi-file-word";
    case "xls":
    case "xlsx":
      return "pi pi-file-excel";
    default:
      return "pi pi-file";
  }
};

// Build a unique id for temporary upload items.
const createTemporaryUploadId = (): string =>
  `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;

const BulkUpload = (): React.ReactElement => {
  const [activeStep, setActiveStep] = useState(1);
  const [loader, setLoader] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [verifyingPaths, setVerifyingPaths] = useState(false);
  const [uploadProgress, setUploadProgress] = useState<IUploadProgress | null>(
    null,
  );
  const [loanFolders, setLoanFolders] = useState<ILoanFolder[]>([]);
  const [pathOptions, setPathOptions] = useState<IPathOption[]>([]);
  const [tagTrees, setTagTrees] = useState<Record<TaxonomyCategory, TreeNode[]>>(
    {} as Record<TaxonomyCategory, TreeNode[]>,
  );
  const [tagFields, setTagFields] = useState<Record<TaxonomyCategory, string>>(
    {} as Record<TaxonomyCategory, string>,
  );
  const [automaticFolderTags, setAutomaticFolderTags] = useState<
    Partial<Record<TaxonomyCategory, ITaxonomyTag>>
  >({});
  const [metadataLoading, setMetadataLoading] = useState(true);
  const [uploadFiles, setUploadFiles] = useState<IUploadFileItem[]>([]);
  const [bulkTagCategory, setBulkTagCategory] =
    useState<TaxonomyCategory | null>(null);
  const [bulkTagTermIds, setBulkTagTermIds] = useState<string[]>([]);
  const [isDragOver, setIsDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isMountedRef = useRef(true);
  const tagFieldsRef = useRef(tagFields);
  const tagTreesRef = useRef(tagTrees);

  tagFieldsRef.current = tagFields;
  tagTreesRef.current = tagTrees;

  const pathDropdownOptions = useMemo(
    () => pathOptions.map((o) => ({ label: o.label, value: o.key })),
    [pathOptions],
  );

  // Load loan folders first so slow taxonomy APIs do not block the whole tab.
  const initializeBulkUploadData = useCallback(async (): Promise<void> => {
    try {
      setLoader(true);
      const list = sp.web.lists.getByTitle(listNames.loan);
      const rootFolder = await withSharePointRetry(() => list.rootFolder());

      const response: any = await withSharePointRetry(() =>
        list.renderListDataAsStream({
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
                <FieldRef Name="FileLeafRef" Ascending="TRUE" />
              </OrderBy>
            </Query>
            <ViewFields>
              <FieldRef Name="FileRef" />
              <FieldRef Name="FileLeafRef" />
            </ViewFields>
            <RowLimit>2000</RowLimit>
          </View>
        `,
          FolderServerRelativeUrl: rootFolder.ServerRelativeUrl,
        }),
      );

      const folders: ILoanFolder[] = (response.Row || [])
        .map((item: any) => ({
          name: String(item.FileLeafRef || "").trim(),
          fileRef: normalizeSharePointPath(String(item.FileRef || "")),
        }))
        .filter(
          (folder: ILoanFolder) =>
            !!folder.name &&
            !!folder.fileRef &&
            !shouldSkipFolderName(folder.name),
        );

      const rootOptions = folders.map((folder) => ({
        key: folder.fileRef,
        label: folder.name,
        loanNumber: folder.name,
        folderPath: folder.fileRef,
      }));
      let options = rootOptions;

      try {
        options = await withSharePointRetry(() =>
          loadFolderPathOptionsFromSharePoint(folders),
        );
      } catch (error) {
        console.error("loadFolderPaths error:", error);
      }

      if (!isMountedRef.current) return;
      setLoanFolders(folders);
      setPathOptions(options);
      setLoader(false);

      // Metadata is needed only after files are mapped, so load it in parallel
      // without keeping the initial page loader visible.
      setMetadataLoading(true);
      const [metadataResult, automaticTagsResult] = await Promise.all([
        loadManagedMetadataTagTreesCached()
          .then((metadata) => ({ metadata, error: null as unknown }))
          .catch((error: unknown) => ({ metadata: null, error })),
        loadAutomaticFolderTagsCached()
          .then((tags) => ({ tags, error: null as unknown }))
          .catch((error: unknown) => ({
            tags: {} as Partial<Record<TaxonomyCategory, ITaxonomyTag>>,
            error,
          })),
      ]);

      if (!isMountedRef.current) return;
      if (metadataResult.metadata) {
        setTagTrees(metadataResult.metadata.trees);
        setTagFields(metadataResult.metadata.fields);
      } else {
        console.error("loadManagedMetadata error:", metadataResult.error);
      }

      setAutomaticFolderTags(automaticTagsResult.tags);
      if (automaticTagsResult.error) {
        console.error(
          "loadAutomaticFolderTags error:",
          automaticTagsResult.error,
        );
      }
      setMetadataLoading(false);
    } catch (error) {
      console.error("loadLoanFolders error:", error);
      if (isMountedRef.current) {
        toastFunc("error", "Error", "Failed to load loan folders");
        setPathOptions([]);
        setMetadataLoading(false);
        setLoader(false);
      }
    } finally {
      if (isMountedRef.current) setLoader(false);
    }
  }, []);

  useEffect(() => {
    isMountedRef.current = true;
    void initializeBulkUploadData();
    return () => {
      isMountedRef.current = false;
    };
  }, [initializeBulkUploadData]);

  // Files
  // Create upload item model from selected file.
  const createUploadQueueItem = (file: File): IUploadFileItem => {
    const pathKey = suggestFolderPathFromFileName(file.name, pathOptions);
    const option = pathOptions.find((o) => o.key === pathKey);
    const folderPath = option?.folderPath || "";
    const tagCategory = determineTagCategoryFromFolderPath(folderPath);
    const automaticTag = tagCategory
      ? automaticFolderTags[tagCategory]
      : undefined;

    return {
      id: createTemporaryUploadId(),
      file,
      fileName: file.name,
      pathKey,
      folderPath,
      tagCategory,
      tags: automaticTag ? [{ ...automaticTag }] : [],
      pathVerified: false,
    };
  };

  // Files
  // Add files with count limits and path auto-suggestions.
  const addFilesToUploadQueue = (files: File[]): void => {
    if (!files.length) return;

    if (!pathOptions.length) {
      toastFunc(
        "warn",
        "Warning",
        "Folder paths are still loading. Please wait and try again.",
      );
      return;
    }

    if (metadataLoading) {
      toastFunc(
        "info",
        "Please wait",
        "Managed metadata is still loading.",
      );
      return;
    }

    const remaining = MAX_FILES - uploadFiles.length;
    if (remaining <= 0) {
      toastFunc("warn", "Warning", `Maximum ${MAX_FILES} files allowed`);
      return;
    }

    const toAdd = files.slice(0, remaining);
    if (files.length > remaining) {
      toastFunc(
        "warn",
        "Warning",
        `Only ${remaining} more file(s) can be added`,
      );
    }

    setUploadFiles((prev) => [
      ...prev,
      ...toAdd.map((file) => createUploadQueueItem(file)),
    ]);
  };

  // Handle file input selection events.
  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files?.length) {
      addFilesToUploadQueue(Array.from(e.target.files));
    }
    e.target.value = "";
  };

  // Handle drag-drop file upload area events.
  const handleDrop = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setIsDragOver(false);
    if (e.dataTransfer.files?.length) {
      addFilesToUploadQueue(Array.from(e.dataTransfer.files));
    }
  };

  // Path
  // Update file path mapping and reset path verification.
  const updateUploadDestination = (id: string, pathKey: string): void => {
    const option = pathOptions.find((o) => o.key === pathKey);
    const folderPath = normalizeSharePointPath(
      option?.folderPath || pathKey || "",
    );
    const tagCategory = determineTagCategoryFromFolderPath(folderPath);
    const automaticTag = tagCategory
      ? automaticFolderTags[tagCategory]
      : undefined;
    setUploadFiles((prev) =>
      prev.map((item) =>
        item.id === id
          ? {
              ...item,
              pathKey: option?.key || folderPath,
              folderPath,
              tagCategory,
              tags: automaticTag ? [{ ...automaticTag }] : [],
              pathVerified: false,
            }
          : item,
      ),
    );
  };

  // Tags
  // Update the selected taxonomy category for one file.
  const updateUploadTagCategory = (
    id: string,
    category: TaxonomyCategory,
  ): void => {
    setUploadFiles((prev) =>
      prev.map((item) =>
        item.id === id
          ? {
              ...item,
              tagCategory: category,
              tags: [],
            }
          : item,
      ),
    );
  };

  // Tags
  // Update selected taxonomy terms for one file item.
  const updateUploadTags = (id: string, termIds: string[]): void => {
    setUploadFiles((prev) =>
      prev.map((item) => {
        if (item.id !== id) return item;
        if (!termIds.length || !item.tagCategory) {
          return { ...item, tags: [] };
        }

        const trees = tagTreesRef.current;
        const fields = tagFieldsRef.current;
        const tags = buildTaxonomyTagPayloads(
          termIds,
          item.tagCategory,
          trees,
          fields,
        );

        return { ...item, tags };
      }),
    );
  };

  // Remove one file from upload queue.
  const removeFileFromUploadQueue = (id: string): void => {
    setUploadFiles((prev) => prev.filter((item) => item.id !== id));
  };

  // Reset wizard state and clear all upload files.
  const clearUploadQueue = (): void => {
    setUploadFiles([]);
    setBulkTagCategory(null);
    setBulkTagTermIds([]);
  };

  // Bulk
  // Apply selected bulk tag to all queued files.
  const applyBulkTagsToAllFiles = (): void => {
    if (!bulkTagCategory || !bulkTagTermIds.length) {
      toastFunc("warn", "Warning", "Select a category and at least one tag");
      return;
    }

    const tags = buildTaxonomyTagPayloads(
      bulkTagTermIds,
      bulkTagCategory,
      tagTrees,
      tagFields,
    );

    if (!tags.length) {
      toastFunc("error", "Error", "Selected tags are invalid");
      return;
    }

    setUploadFiles((prev) =>
      prev.map((item) => ({
        ...item,
        tagCategory: bulkTagCategory,
        tags: [...tags],
      })),
    );

    toastFunc(
      "success",
      "Success",
      `${tags.length} tag(s) applied to ${uploadFiles.length} file(s).`,
    );
  };

  // Verify
  // Verify that a target folder path exists in SharePoint.
  const doesFolderPathExist = async (folderPath: string): Promise<boolean> => {
    if (!folderPath) return false;
    try {
      await sp.web
        .getFolderByServerRelativePath(folderPath)
        .select("ServerRelativeUrl")();
      return true;
    } catch {
      return false;
    }
  };

  // Verify
  // Verify all selected folder paths before submit.
  const verifyAllFolderPaths = async (): Promise<boolean> => {
    const uniquePaths = Array.from(
      new Set(
        uploadFiles
          .map((item) => item.folderPath)
          .filter((path) => !!path),
      ),
    );

    const pathResults = await Promise.all(
      uniquePaths.map(async (path) => ({
        path,
        verified: await doesFolderPathExist(path),
      })),
    );

    const verifiedMap = new Map(
      pathResults.map((result) => [result.path, result.verified]),
    );

    setUploadFiles((prev) =>
      prev.map((item) => ({
        ...item,
        pathVerified: !!verifiedMap.get(item.folderPath),
      })),
    );

    return pathResults.every((result) => result.verified);
  };

  // Validate
  // Validate step one inputs before opening tags step.
  const validateFileMappingStep = (): boolean => {
    if (!uploadFiles.length) {
      toastFunc("warn", "Warning", "Add at least one file to continue");
      return false;
    }

    const missingPath = uploadFiles.some(
      (file) => !file.pathKey || !file.folderPath,
    );
    if (missingPath) {
      toastFunc("warn", "Warning", "Select a path for every file");
      return false;
    }

    return true;
  };

  // Validate
  // Validate tag assignments before review step.
  const validateTagAssignmentStep = (): boolean => {
    const missingTag = uploadFiles.some(
      (file) => file.tagCategory && !file.tags.length,
    );
    if (missingTag) {
      toastFunc(
        "warn",
        "Warning",
        "Select at least one tag for each file that has a category assigned",
      );
      return false;
    }
    return true;
  };

  // Steps
  // Move from mapping step to tag assignment step.
  const navigateToTagAssignmentStep = (): void => {
    if (!validateFileMappingStep()) return;
    setActiveStep(2);
  };

  // Steps
  // Move from tag step to review step after validation.
  const navigateToReviewStep = async (): Promise<void> => {
    if (!validateTagAssignmentStep() || verifyingPaths) return;

    setVerifyingPaths(true);
    try {
      const allValid = await verifyAllFolderPaths();

      if (!allValid) {
        toastFunc(
          "warn",
          "Warning",
          "One or more folder paths do not exist in SharePoint",
        );
        return;
      }

      setActiveStep(3);
    } finally {
      setVerifyingPaths(false);
    }
  };

  // Submit
  // Upload files in parallel and apply managed metadata to uploaded items.
  const handleSubmit = async (): Promise<void> => {
    if (!uploadFiles.length || submitting) return;

    setSubmitting(true);
    setUploadProgress({
      total: uploadFiles.length,
      completed: 0,
      successCount: 0,
      failCount: 0,
      currentFileName: uploadFiles[0]?.fileName || "",
      phase: "uploading",
    });

    try {
      const results = await runTasksWithConcurrency(
        uploadFiles,
        UPLOAD_CONCURRENCY,
        async (item) => {
          setUploadProgress((prev) =>
            prev
              ? {
                  ...prev,
                  currentFileName: item.fileName,
                }
              : null,
          );

          const result = await uploadFileWithMetadata(item);

          setUploadProgress((prev) =>
            prev
              ? {
                  ...prev,
                  completed: prev.completed + 1,
                  successCount: prev.successCount + (result.success ? 1 : 0),
                  failCount: prev.failCount + (result.success ? 0 : 1),
                  currentFileName: item.fileName,
                }
              : null,
          );

          return result;
        },
      );

      const successCount = results.filter((result) => result.success).length;
      const failCount = results.length - successCount;

      setUploadProgress((prev) =>
        prev ? { ...prev, phase: "complete" } : null,
      );

      if (successCount > 0) {
        toastFunc(
          "success",
          "Success",
          `${successCount} file(s) uploaded successfully`,
        );
      }

      if (failCount > 0) {
        const failedNames = results
          .filter((result) => !result.success)
          .slice(0, 3)
          .map((result) => result.fileName)
          .join(", ");
        const suffix =
          failCount > 3 ? ` and ${failCount - 3} more` : "";

        toastFunc(
          "error",
          "Error",
          `${failCount} file(s) failed to upload${failedNames ? `: ${failedNames}${suffix}` : ""}`,
        );
      }

      if (successCount > 0 && failCount === 0) {
        await new Promise((resolve) => setTimeout(resolve, 800));
        clearUploadQueue();
        setActiveStep(1);
      }
    } finally {
      setSubmitting(false);
      setUploadProgress(null);
    }
  };

  // Cancel the workflow and reset wizard to first step.
  const handleCancel = (): void => {
    clearUploadQueue();
    setActiveStep(1);
  };

  // Resolve display label for selected folder path.
  const getFolderPathDisplayLabel = (folderPath: string): string =>
    pathOptions.find((option) =>
      areSharePointPathsEqual(option.key, folderPath),
    )?.label ||
    formatFolderPathLabel(folderPath, loanFolders);

  // Render one step indicator in the top progress bar.
  const renderStepIndicator = (
    step: number,
    label: string,
  ): React.ReactElement => {
    const isCompleted = activeStep > step;
    const isActive = activeStep === step;

    return (
      <div
        className={`${styles.step} ${isActive ? styles.active : ""} ${
          isCompleted ? styles.completed : ""
        }`}
      >
        <div className={styles.circle}>
          {isCompleted ? <i className="pi pi-check" /> : step}
        </div>
        <span>{label}</span>
      </div>
    );
  };

  const allPathsVerified = uploadFiles.every((f) => f.pathVerified);
  const allTagsAssigned = uploadFiles.every(
    (file) => !file.tagCategory || file.tags.length > 0,
  );

  const renderSubmitProgress = (): React.ReactElement | null => {
    if (!uploadProgress) return null;

    const percent = uploadProgress.total
      ? Math.round((uploadProgress.completed / uploadProgress.total) * 100)
      : 0;
    const isComplete = uploadProgress.phase === "complete";

    return (
      <div className={styles.submitOverlay} role="dialog" aria-modal="true">
        <div className={styles.submitProgressCard}>
          <div className={styles.submitProgressHeader}>
            <div
              className={`${styles.submitProgressSpinner} ${
                isComplete ? styles.submitProgressSpinnerDone : ""
              }`}
            >
              {isComplete ? (
                <i className="pi pi-check" />
              ) : (
                <i className="pi pi-cloud-upload" />
              )}
            </div>
            <div>
              <h3 className={styles.submitProgressTitle}>
                {isComplete ? "Upload complete" : "Uploading files"}
              </h3>
              <p className={styles.submitProgressSubtitle}>
                {isComplete
                  ? `${uploadProgress.successCount} of ${uploadProgress.total} files uploaded successfully`
                  : `Processing ${uploadProgress.completed} of ${uploadProgress.total} files`}
              </p>
            </div>
          </div>

          <div className={styles.submitProgressBarTrack}>
            <div
              className={styles.submitProgressBarFill}
              style={{ width: `${percent}%` }}
            />
          </div>

          <div className={styles.submitProgressMeta}>
            <span>{percent}%</span>
            <span>
              {uploadProgress.successCount} succeeded
              {uploadProgress.failCount > 0
                ? ` • ${uploadProgress.failCount} failed`
                : ""}
            </span>
          </div>

          {!isComplete && uploadProgress.currentFileName && (
            <div className={styles.submitProgressCurrent}>
              <span className={styles.submitProgressCurrentLabel}>
                Current file
              </span>
              <span
                className={styles.submitProgressCurrentName}
                title={uploadProgress.currentFileName}
              >
                {uploadProgress.currentFileName}
              </span>
            </div>
          )}

          <p className={styles.submitProgressHint}>
            {isComplete
              ? "Finishing up..."
              : "Please keep this page open until the upload finishes."}
          </p>
        </div>
      </div>
    );
  };

  if (loader && !uploadFiles.length && activeStep === 1) {
    return <Loader />;
  }

  return (
    <div className={styles.bulkUploadWrapper}>
      {renderSubmitProgress()}

      <div className={styles.stepper}>
        {renderStepIndicator(1, "Select & Map Files")}
        <div className={styles.connector} />
        {renderStepIndicator(2, "Assign Tags")}
        <div className={styles.connector} />
        {renderStepIndicator(3, "Review & Submit")}
      </div>

      {/* ─── Step 1: Select & Map Files ─────────────────────────────────── */}
      {activeStep === 1 && (
        <div className={styles.stepContent}>
          <div
            className={`${styles.uploadArea} ${isDragOver ? styles.dragOver : ""}`}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragOver(true);
            }}
            onDragLeave={() => setIsDragOver(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            role="button"
            tabIndex={0}
            onKeyDown={(e) => {
              if (e.key === "Enter" || e.key === " ") {
                fileInputRef.current?.click();
              }
            }}
          >
            <i className={`pi pi-upload ${styles.uploadIcon}`} />
            <div className={styles.uploadTitle}>
              Click or drag files here to upload
            </div>
            <div className={styles.uploadText}>
              Supports up to {MAX_FILES} files — path auto-suggested from
              filename prefix
            </div>
            <input
              ref={fileInputRef}
              type="file"
              multiple
              className={styles.hiddenInput}
              onChange={handleFileSelect}
            />
          </div>

          {uploadFiles.length > 0 && (
            <div className={styles.fileGrid}>
              {uploadFiles.map((item) => (
                <div key={item.id} className={styles.fileCard}>
                  <button
                    type="button"
                    className={styles.removeBtn}
                    onClick={() => removeFileFromUploadQueue(item.id)}
                    aria-label="Remove file"
                  >
                    <i className="pi pi-times" />
                  </button>
                  <div className={styles.fileCardHeader}>
                    <i
                      className={`${getFileIconClassName(item.fileName)} ${styles.fileTypeIcon}`}
                    />
                    <span className={styles.fileName} title={item.fileName}>
                      {item.fileName}
                    </span>
                  </div>
                  <label className={styles.fieldLabel}>Path / Location</label>
                  <Dropdown
                    value={item.pathKey}
                    options={pathDropdownOptions}
                    optionLabel="label"
                    optionValue="value"
                    onChange={(e) =>
                      updateUploadDestination(item.id, e.value)
                    }
                    placeholder={
                      pathDropdownOptions.length
                        ? "Select folder path"
                        : "No folders available"
                    }
                    className={styles.pathDropdown}
                    filter
                    filterPlaceholder="Search path..."
                    appendTo={document.body}
                    disabled={!pathDropdownOptions.length}
                  />
                </div>
              ))}
            </div>
          )}

          <div className={styles.stepFooter}>
            {uploadFiles.length > 0 && (
              <Button
                label="Clear all"
                outlined
                className={styles.clearBtn}
                onClick={clearUploadQueue}
              />
            )}
            <Button
              label="Next: Assign Tags →"
              className={styles.primaryBtn}
              onClick={navigateToTagAssignmentStep}
              disabled={!uploadFiles.length}
            />
          </div>
        </div>
      )}

      {/* ─── Step 2: Assign Tags ────────────────────────────────────────── */}
      {activeStep === 2 && (
        <div className={styles.stepContent}>
          <div className={styles.bulkTagBar}>
            <span className={styles.bulkTagLabel}>Apply tags:</span>
            <div className={styles.tagPills}>
              {TAXONOMY_CATEGORIES.map((tag) => (
                <button
                  key={tag}
                  type="button"
                  className={`${styles.tagPill} ${
                    bulkTagCategory === tag ? styles.tagPillActive : ""
                  }`}
                  data-active={bulkTagCategory === tag ? "true" : "false"}
                  aria-pressed={bulkTagCategory === tag}
                  onClick={() => {
                    setBulkTagCategory(tag);
                    setBulkTagTermIds([]);
                  }}
                >
                  {tag}
                </button>
              ))}
            </div>
            <div className={styles.bulkTagPicker}>
              <TaxonomyTagPicker
                key={`bulk-${bulkTagCategory || "none"}`}
                category={bulkTagCategory}
                tree={bulkTagCategory ? tagTrees[bulkTagCategory] || [] : []}
                value={bulkTagTermIds}
                placeholder="— select managed metadata tags —"
                disabled={!bulkTagCategory}
                onChange={setBulkTagTermIds}
              />
            </div>
            <Button
              label="Apply to all files"
              outlined
              className={styles.applyAllBtn}
              onClick={applyBulkTagsToAllFiles}
            />
          </div>

          <p className={styles.sectionHint}>Or assign tags per file:</p>

          <div className={styles.fileGrid}>
            {uploadFiles.map((item) => (
              <div key={item.id} className={styles.fileCard}>
                <div className={styles.fileCardHeader}>
                  <i
                    className={`${getFileIconClassName(item.fileName)} ${styles.fileTypeIcon}`}
                  />
                  <span className={styles.fileName} title={item.fileName}>
                    {item.fileName}
                  </span>
                </div>
                {item.folderPath && (
                  <div className={styles.reviewPath} title={item.folderPath}>
                    {getFolderPathDisplayLabel(item.folderPath)}
                  </div>
                )}
                <label className={styles.fieldLabel}>Tag category</label>
                <div className={styles.tagPills}>
                  {TAXONOMY_CATEGORIES.map((tag) => (
                    <button
                      key={`${item.id}-${tag}`}
                      type="button"
                      className={`${styles.tagPill} ${
                        item.tagCategory === tag ? styles.tagPillActive : ""
                      }`}
                      data-active={item.tagCategory === tag ? "true" : "false"}
                      aria-pressed={item.tagCategory === tag}
                      onClick={() => updateUploadTagCategory(item.id, tag)}
                    >
                      {tag}
                    </button>
                  ))}
                </div>
                <label className={styles.fieldLabel}>
                  Tag
                  {item.tagCategory ? ` (${item.tagCategory})` : ""}
                </label>
                <TaxonomyTagPicker
                  key={`${item.id}-${item.tagCategory || "none"}`}
                  category={item.tagCategory}
                  tree={
                    item.tagCategory ? tagTrees[item.tagCategory] || [] : []
                  }
                  value={item.tags.map((tag) => tag.termId)}
                  displayLabels={item.tags.map((tag) => tag.label)}
                  onChange={(termIds) => updateUploadTags(item.id, termIds)}
                />
              </div>
            ))}
          </div>

          <div className={styles.stepFooter}>
            <Button
              label="← Back"
              outlined
              className={styles.backBtn}
              onClick={() => setActiveStep(1)}
            />
            <Button
              label={verifyingPaths ? "Verifying paths..." : "Next: Review →"}
              className={styles.primaryBtn}
              icon={verifyingPaths ? "pi pi-spin pi-spinner" : undefined}
              onClick={() => void navigateToReviewStep()}
              disabled={verifyingPaths}
            />
          </div>
        </div>
      )}

      {/* ─── Step 3: Review & Submit ────────────────────────────────────── */}
      {activeStep === 3 && (
        <div className={styles.stepContent}>
          <h3 className={styles.reviewTitle}>
            Review {uploadFiles.length} file
            {uploadFiles.length !== 1 ? "s" : ""} before submitting
          </h3>

          <div className={styles.reviewPanel}>
            <div className={styles.reviewPanelHeader}>
              <span>{uploadFiles.length} files ready to upload</span>
              <span className={styles.reviewStatus}>
                {allPathsVerified ? "All paths verified" : "Paths pending"}
                {" • "}
                {allTagsAssigned ? "Tags assigned" : "Tags pending"}
              </span>
            </div>

            <div className={styles.reviewList}>
              {uploadFiles.map((item) => (
                <div key={item.id} className={styles.reviewRow}>
                  <div className={styles.reviewFileInfo}>
                    <i
                      className={`${getFileIconClassName(item.fileName)} ${styles.fileTypeIcon}`}
                    />
                    <span title={item.fileName}>{item.fileName}</span>
                  </div>
                  <div className={styles.reviewMeta}>
                    <span className={styles.reviewPath}>
                      {getFolderPathDisplayLabel(item.pathKey)}
                    </span>
                    {item.tags.length > 0 && (
                      <span className={styles.reviewTag}>
                        {item.tags.map((tag) => tag.label).join(", ")}
                      </span>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div className={styles.stepFooter}>
            <Button
              label="← Back"
              outlined
              className={styles.backBtn}
              onClick={() => setActiveStep(2)}
            />
            <div className={styles.reviewActions}>
              <Button
                label="Cancel"
                outlined
                className={styles.cancelBtn}
                onClick={handleCancel}
              />
              <Button
                label={
                  submitting
                    ? `Uploading ${uploadProgress?.completed || 0} of ${uploadFiles.length}...`
                    : `Submit ${uploadFiles.length} File${
                        uploadFiles.length !== 1 ? "s" : ""
                      }`
                }
                icon={submitting ? "pi pi-spin pi-spinner" : "pi pi-check"}
                className={styles.primaryBtn}
                onClick={() => void handleSubmit()}
                disabled={submitting}
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default BulkUpload;
