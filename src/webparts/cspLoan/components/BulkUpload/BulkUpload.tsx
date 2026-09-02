/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from "react";
import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { sp } from "@pnp/sp/presets/all";
import { Button } from "primereact/button";
import { Dropdown } from "primereact/dropdown";
import { Tree, TreeSelectionEvent } from "primereact/tree";
import { TreeNode } from "primereact/treenode";
import { OverlayPanel } from "primereact/overlaypanel";
import styles from "./BulkUpload.module.scss";
import {
  bulkUploadConfig,
  folderStructure,
  listNames,
  managedMetadataFields,
  loanLibraryFields,
  parseLoanSponsorLookup,
  buildSponsorLookupUpdatePayload,
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

const ensureServerRelativePath = (path: string): string => {
  const normalized = normalizeSharePointPath(path);
  if (!normalized) return "";
  return normalized.startsWith("/") ? normalized : `/${normalized}`;
};

const isSharePointFolderNotFoundError = (error: unknown): boolean => {
  const message = (
    error instanceof Error ? error.message : String(error || "")
  ).toLowerCase();

  return (
    message.includes("does not exist") ||
    message.includes("not found") ||
    message.includes("404") ||
    message.includes("file not found") ||
    message.includes("cannot find") ||
    message.includes("unable to find") ||
    message.includes("item does not exist") ||
    message.includes("folder not found")
  );
};

const areSharePointPathsEqual = (left: string, right: string): boolean =>
  normalizeSharePointPath(left).toLowerCase() ===
  normalizeSharePointPath(right).toLowerCase();

const areFileNamesEqual = (left: string, right: string): boolean =>
  left.trim().toLowerCase() === right.trim().toLowerCase();

// Same file name mapped to the same folder path more than once.
const isSameFileSamePath = (
  left: Pick<IUploadFileItem, "fileName" | "folderPath" | "pathKey">,
  right: Pick<IUploadFileItem, "fileName" | "folderPath" | "pathKey">,
): boolean => {
  if (!areFileNamesEqual(left.fileName, right.fileName)) return false;
  const leftPath = left.folderPath || left.pathKey;
  const rightPath = right.folderPath || right.pathKey;
  if (!leftPath || !rightPath) return false;
  return areSharePointPathsEqual(leftPath, rightPath);
};

const getDuplicateSamePathFileIds = (files: IUploadFileItem[]): string[] => {
  const groups = new Map<string, string[]>();

  files.forEach((file) => {
    const path = file.folderPath || file.pathKey;
    if (!path) return;
    const key = `${file.fileName.trim().toLowerCase()}::${normalizeSharePointPath(path).toLowerCase()}`;
    const ids = groups.get(key) || [];
    ids.push(file.id);
    groups.set(key, ids);
  });

  const duplicateIds: string[] = [];
  groups.forEach((ids) => {
    if (ids.length > 1) duplicateIds.push(...ids);
  });
  return duplicateIds;
};

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
      sponsorId: loanFolder.sponsorId ?? null,
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
      sponsorId: loanFolder.sponsorId ?? null,
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

// Auto-map folder categories to matching terms from each category's own MMS field.
// Uses Asset Management / Legal / Servicing fields (no shared "Tags" column required).
const loadAutomaticFolderTags = async (): Promise<
  Partial<Record<TaxonomyCategory, ITaxonomyTag>>
> => {
  const { trees, fields } = await loadManagedMetadataTagTreesCached();
  const tags: Partial<Record<TaxonomyCategory, ITaxonomyTag>> = {};

  for (const category of TAXONOMY_CATEGORIES) {
    const fieldInternalName = fields[category];
    if (!fieldInternalName) continue;

    const term = findTermNodeByLabel(trees[category] || [], category);
    const termId = String(term?.key || "");
    if (!term || !termId) continue;

    tags[category] = {
      termId,
      label: String(term.data?.label || term.label || category),
      fieldInternalName,
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

// Upload one file and apply sponsor + managed metadata when available.
const uploadFileWithMetadata = async (
  item: IUploadFileItem,
): Promise<IUploadResult> => {
  try {
    const result = await sp.web
      .getFolderByServerRelativePath(ensureServerRelativePath(item.folderPath))
      .files.add(item.fileName, item.file, true);

    const needsItemUpdate =
      (Number(item.sponsorId) || 0) > 0 || item.tags.length > 0;

    if (needsItemUpdate) {
      const fileItem: any = await result.file.listItemAllFields();
      const itemId = Number(fileItem.Id ?? fileItem.ID);

      if (!itemId) {
        throw new Error(`Could not resolve list item ID for ${item.fileName}`);
      }

      const sponsorId = Number(item.sponsorId) || 0;
      if (sponsorId > 0) {
        await sp.web.lists
          .getByTitle(listNames.loan)
          .items.getById(itemId)
          .update(buildSponsorLookupUpdatePayload(sponsorId));
      }

      if (item.tags.length) {
        await applyTaxonomyTagsToListItem(itemId, item.tags);
      }
    }

    return { fileName: item.fileName, success: true };
  } catch (error) {
    const folderNotFound = isSharePointFolderNotFoundError(error);
    return {
      fileName: item.fileName,
      success: false,
      folderNotFound,
      folderPath: item.folderPath,
      error: folderNotFound
        ? "Folder does not exist in SharePoint library"
        : error instanceof Error
          ? error.message
          : "Upload failed",
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
            findTermLabelById(tree, termId) ||
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

  // Close panel when category is cleared / picker disabled so UI fully resets.
  useEffect(() => {
    if (disabled || !category) {
      overlayRef.current?.hide();
      setOverlayVisible(false);
      if (!value.length) {
        syncSelectionFromValue([]);
      }
    }
  }, [category, disabled, value.length, syncSelectionFromValue]);

  useEffect(() => {
    if (!overlayVisible) return;
    const onDocumentMouseDown = (event: MouseEvent): void => {
      const target = event.target as HTMLElement | null;
      if (!target) return;

      if (
        triggerRef.current?.contains(target) ||
        target.closest?.(".p-overlaypanel") ||
        target.closest?.(`.${styles.taxonomyOverlay}`)
      ) {
        return;
      }

      overlayRef.current?.hide();
      setOverlayVisible(false);
    };
    document.addEventListener("mousedown", onDocumentMouseDown);
    return () => document.removeEventListener("mousedown", onDocumentMouseDown);
  }, [overlayVisible]);

  // Prefer controlled `value` so clearing parent state never leaves stale labels.
  const resolvedLabel =
    value.length === 0
      ? ""
      : value
          .map((termId) => findTermLabelById(tree, termId))
          .filter(Boolean)
          .join(", ") ||
        (displayLabelsProp || []).join(", ") ||
        selectedLabels.join(", ");
  const hasSelection = value.length > 0;

  const openOverlay = (event: React.MouseEvent<HTMLButtonElement>): void => {
    if (disabled || !category) return;
    overlayRef.current?.show(event, event.currentTarget);
    setOverlayVisible(true);
  };

  const closeOverlay = (): void => {
    overlayRef.current?.hide();
    setOverlayVisible(false);
  };

  const toggleOverlay = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    event.stopPropagation();
    if (overlayVisible) {
      closeOverlay();
      return;
    }
    openOverlay(event);
  };

  const handleToggle = (termId: string, label: string): void => {
    const currentTermIds = selectedTermIdsRef.current;
    const currentLabels = selectedLabelsRef.current;
    const isSelected = currentTermIds.some((id) =>
      areTermIdsEqual(id, termId),
    );
    const nextIds = isSelected
      ? currentTermIds.filter((id) => !areTermIdsEqual(id, termId))
      : [...currentTermIds, termId];

    syncSelectionFromValue(
      nextIds,
      isSelected
        ? currentLabels.filter(
            (_, index) => !areTermIdsEqual(currentTermIds[index], termId),
          )
        : [...currentLabels, label],
    );
    onChange(nextIds);
  };

  const nodeTemplate = (node: TreeNode): React.ReactElement => {
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
        <i className={`pi pi-tag ${styles.copyMoveTagIcon}`} />
        <span className={styles.copyMoveNodeLabel} title={label}>
          {label}
        </span>
        {isSelected && (
          <i
            className={`pi pi-times ${styles.taxonomyTagRemoveIcon}`}
            title="Remove tag"
            aria-label={`Remove ${label}`}
          />
        )}
      </span>
    );
  };

  return (
    <>
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
            hasSelection ? styles.taxonomyTagValue : styles.pathPlaceholder
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
          } ${styles.pathPickerIcon}`}
        />
      </button>

      <OverlayPanel
        ref={overlayRef}
        className={styles.taxonomyOverlay}
        dismissable={false}
        onHide={() => setOverlayVisible(false)}
      >
        <div className={styles.taxonomyOverlayTitle}>
          <span>{category ? `${category} tags` : "Tags"}</span>
          <span className={styles.taxonomyOverlayHint}>
            Click again to remove
          </span>
        </div>
        <div className={styles.taxonomyTreePanel}>
          {tree.length ? (
            <Tree
              className={styles.copyMoveTree}
              value={tree}
              nodeTemplate={nodeTemplate}
              propagateSelectionUp={false}
              propagateSelectionDown={false}
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

interface IPathDropdownOption {
  label: string;
  value: string;
}

const renderPathDropdownValue = (
  selectedOption: IPathDropdownOption | null | undefined,
): React.ReactNode => {
  if (!selectedOption) {
    return (
      <span className={styles.pathDropdownPlaceholder}>Select folder path</span>
    );
  }

  return (
    <span className={styles.pathDropdownValue} title={selectedOption.label}>
      {selectedOption.label}
    </span>
  );
};

const renderPathDropdownItem = (option: IPathDropdownOption): React.ReactNode => (
  <span className={styles.pathDropdownItem} title={option.label}>
    {option.label}
  </span>
);

interface IUploadPathDropdownProps {
  value: string;
  options: IPathDropdownOption[];
  hasError: boolean;
  disabled: boolean;
  onChange: (pathKey: string) => void;
}

const UploadPathDropdown = ({
  value,
  options,
  hasError,
  disabled,
  onChange,
}: IUploadPathDropdownProps): React.ReactElement => {
  const wrapRef = useRef<HTMLDivElement>(null);
  const [isPanelOpen, setIsPanelOpen] = useState(false);
  const [panelStyle, setPanelStyle] = useState<React.CSSProperties>({
    visibility: "hidden",
    opacity: 0,
  });

  const getTriggerElement = useCallback((): HTMLElement | null => {
    return (
      (wrapRef.current?.querySelector(".p-dropdown") as HTMLElement | null) ||
      wrapRef.current
    );
  }, []);

  const buildPanelLayout = useCallback((): React.CSSProperties | null => {
    const trigger = getTriggerElement();
    if (!trigger) {
      return null;
    }

    const rect = trigger.getBoundingClientRect();
    const width = Math.round(rect.width);

    return {
      boxSizing: "border-box",
      width: `${width}px`,
      maxWidth: `${width}px`,
      minWidth: `${width}px`,
      transform: "none",
      margin: 0,
      position: "fixed",
      left: "auto",
      right: `${Math.round(window.innerWidth - rect.right)}px`,
      top: `${Math.round(rect.bottom + 4)}px`,
      visibility: "visible",
      opacity: 1,
    };
  }, [getTriggerElement]);

  const applyPanelLayoutToDom = useCallback(
    (visible: boolean): void => {
      const layout = buildPanelLayout();
      const panel = document.querySelector(
        `.${styles.pathDropdownPanel}`,
      ) as HTMLElement | null;

      if (!layout || !panel) {
        return;
      }

      Object.assign(panel.style, layout);
      panel.style.visibility = visible ? "visible" : "hidden";
      panel.style.opacity = visible ? "1" : "0";
      panel.classList.toggle(styles.pathDropdownPanelReady, visible);
    },
    [buildPanelLayout],
  );

  useLayoutEffect(() => {
    if (!isPanelOpen) {
      return;
    }

    const layout = buildPanelLayout();
    if (layout) {
      setPanelStyle(layout);
    }

    applyPanelLayoutToDom(true);
  }, [isPanelOpen, applyPanelLayoutToDom, buildPanelLayout]);

  const handlePointerDown = (): void => {
    const layout = buildPanelLayout();
    if (layout) {
      setPanelStyle({
        ...layout,
        visibility: "hidden",
        opacity: 0,
      });
    }
  };

  const handleShow = (): void => {
    const layout = buildPanelLayout();
    if (layout) {
      setPanelStyle(layout);
    }
    applyPanelLayoutToDom(true);
    setIsPanelOpen(true);
  };

  const handleHide = (): void => {
    setIsPanelOpen(false);
    setPanelStyle({
      visibility: "hidden",
      opacity: 0,
    });

    const panel = document.querySelector(
      `.${styles.pathDropdownPanel}`,
    ) as HTMLElement | null;
    panel?.classList.remove(styles.pathDropdownPanelReady);
  };

  return (
    <div
      ref={wrapRef}
      className={styles.pathDropdownWrap}
      onPointerDown={handlePointerDown}
    >
      <Dropdown
        value={value}
        options={options}
        optionLabel="label"
        optionValue="value"
        onChange={(e) => onChange(String(e.value || ""))}
        placeholder={
          options.length ? "Select folder path" : "No folders available"
        }
        className={`${styles.pathDropdown} ${
          hasError ? styles.pathDropdownError : ""
        }`}
        panelClassName={styles.pathDropdownPanel}
        panelStyle={panelStyle}
        valueTemplate={renderPathDropdownValue}
        itemTemplate={renderPathDropdownItem}
        onShow={handleShow}
        onHide={handleHide}
        filter
        filterPlaceholder="Search path..."
        scrollHeight="240px"
        appendTo={document.body}
        disabled={disabled || !options.length}
      />
    </div>
  );
};

// Build a unique id for temporary upload items.
const createTemporaryUploadId = (): string =>
  `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;

interface IFolderStructureNode {
  name: string;
  children?: IFolderStructureNode[];
}

type BulkUploadPathMode = "same" | "different";

const PATH_MODE_OPTIONS: {
  value: BulkUploadPathMode;
  title: string;
  description: string;
  example: string;
  icon: string;
}[] = [
  {
    value: "same",
    title: "Same Path",
    description:
      "Upload multiple files to the same folder structure. Loan number is read from each file name.",
    example: "*/Asset Management/Underwriting",
    icon: "pi pi-copy",
  },
  {
    value: "different",
    title: "Different Path",
    description:
      "Map each file to its own destination folder. Paths are auto-suggested from file names.",
    example: "Per-file folder mapping",
    icon: "pi pi-sitemap",
  },
];

const buildFolderStructureTreeNodes = (
  nodes: IFolderStructureNode[],
  parentPath = "",
): TreeNode[] =>
  nodes.map((node) => {
    const relativePath = parentPath ? `${parentPath}/${node.name}` : node.name;
    const childNodes = node.children?.length
      ? buildFolderStructureTreeNodes(node.children, relativePath)
      : undefined;

    return {
      key: relativePath,
      label: node.name,
      data: { relativePath },
      children: childNodes,
      icon: childNodes?.length ? "pi pi-folder" : "pi pi-folder-open",
    };
  });

const findFolderStructureNode = (
  nodes: TreeNode[],
  key: string,
): TreeNode | undefined => {
  for (const node of nodes) {
    if (String(node.key) === key) {
      return node;
    }

    if (node.children?.length) {
      const match = findFolderStructureNode(node.children, key);
      if (match) {
        return match;
      }
    }
  }

  return undefined;
};

const SamePathFolderNodeTemplate = ({
  node,
  expanded,
  isSelected,
}: {
  node: TreeNode;
  expanded: boolean;
  isSelected: boolean;
}): React.ReactElement => (
  <span
    className={`${styles.samePathFolderNode} ${
      isSelected ? styles.copyMoveNodeSelected : ""
    }`}
  >
    <i
      className={`pi ${
        expanded ? "pi-folder-open" : "pi-folder"
      } ${styles.copyMoveTagIcon}`}
    />
    <span className={styles.copyMoveNodeLabel} title={String(node.label || "")}>
      {node.label}
    </span>
  </span>
);

const buildSamePathPrefixedFileName = (
  loanNumber: string,
  fileName: string,
): string => {
  const prefix = `${loanNumber}_`;
  return fileName.startsWith(prefix) ? fileName : `${prefix}${fileName}`;
};

const parseLoanNumberFromFileName = (
  fileName: string,
  loanFolders: ILoanFolder[],
): string => {
  const stem = fileName.replace(/\.[^/.]+$/, "");
  const underscoreMatch = stem.match(/^([a-z0-9]+)_/i);
  if (underscoreMatch) {
    const candidate = underscoreMatch[1];
    const matched = loanFolders.find(
      (folder) =>
        normalizeTextForMatching(folder.name) ===
        normalizeTextForMatching(candidate),
    );
    if (matched) return matched.name;
  }

  const normalizedStem = normalizeTextForMatching(stem);
  const loanNumbers = loanFolders
    .map((folder) => folder.name)
    .sort((left, right) => right.length - left.length);

  for (const loanNumber of loanNumbers) {
    if (normalizedStem.startsWith(normalizeTextForMatching(loanNumber))) {
      return loanNumber;
    }
  }

  return "";
};

const formatSamePathDisplayLabel = (
  loanNumber: string,
  structureRelativePath: string,
): string => `${loanNumber}/${structureRelativePath}`;

const formatSamePathFromFolderPath = (
  folderPath: string,
  loanFolders: ILoanFolder[],
): string => {
  if (!folderPath) return "";
  const normalized = normalizeSharePointPath(folderPath);
  const loan = loanFolders.find(
    (folder) =>
      areSharePointPathsEqual(normalized, folder.fileRef) ||
      normalized.startsWith(`${normalizeSharePointPath(folder.fileRef)}/`),
  );

  if (!loan) return formatFolderPathLabel(folderPath, loanFolders);

  const relative = normalized
    .slice(normalizeSharePointPath(loan.fileRef).length)
    .replace(/^\//, "");
  return relative ? `${loan.name}/${relative}` : loan.name;
};

type SamePathFileIssue =
  | "missing_loan_number"
  | "missing_folder_selection"
  | "unresolved_path"
  | "folder_not_in_sharepoint";

const getSamePathFileIssue = (
  item: IUploadFileItem,
  loanFolders: ILoanFolder[],
  samePathFolderKey: string,
  invalidFolderPathFileIds: string[] = [],
): SamePathFileIssue | null => {
  if (invalidFolderPathFileIds.includes(item.id)) {
    return "folder_not_in_sharepoint";
  }

  const loanNumber = parseLoanNumberFromFileName(item.file.name, loanFolders);
  if (!loanNumber) return "missing_loan_number";
  if (!samePathFolderKey) return "missing_folder_selection";
  if (!item.folderPath) return "unresolved_path";
  return null;
};

const getSamePathFileIssueMessage = (
  issue: SamePathFileIssue,
  loanNumber: string,
  samePathFolderKey: string,
  folderPath = "",
  loanFolders: ILoanFolder[] = [],
): string => {
  switch (issue) {
    case "missing_loan_number":
      return "File name must use format: {loanNumber}_{fileName} (e.g. 3000101_Appraisal.pdf)";
    case "missing_folder_selection":
      return "Select a folder path from the folder structure (e.g. Asset Management → Underwriting).";
    case "unresolved_path":
      return loanNumber && samePathFolderKey
        ? `Folder path not found: ${formatSamePathDisplayLabel(loanNumber, samePathFolderKey)}. Verify the loan folder exists in SharePoint.`
        : "Could not resolve folder path for this file.";
    case "folder_not_in_sharepoint":
      return folderPath
        ? `Folder does not exist: ${formatSamePathFromFolderPath(folderPath, loanFolders)}. This path is not in the SharePoint library. Please create the folder and try again.`
        : "Folder does not exist. This path is not in the SharePoint library. Please create the folder and try again.";
    default:
      return "Folder path could not be resolved.";
  }
};

const resolveSamePathDestination = (
  loanNumber: string,
  structureRelativePath: string,
  folders: ILoanFolder[],
): { folderPath: string; pathKey: string; sponsorId: number | null } | null => {
  const loanFolder = folders.find((folder) => folder.name === loanNumber);
  if (!loanFolder || !structureRelativePath) return null;

  const folderPath = normalizeSharePointPath(
    `${loanFolder.fileRef}/${structureRelativePath}`,
  );

  return {
    folderPath,
    pathKey: folderPath,
    sponsorId: loanFolder.sponsorId ?? null,
  };
};

const createSamePathUploadQueueItem = (
  file: File,
  structureRelativePath: string,
  folders: ILoanFolder[],
  automaticFolderTags: Partial<Record<TaxonomyCategory, ITaxonomyTag>>,
): IUploadFileItem => {
  const loanNumber = parseLoanNumberFromFileName(file.name, folders);
  const prefixedName = loanNumber
    ? buildSamePathPrefixedFileName(loanNumber, file.name)
    : file.name;

  if (!structureRelativePath || !loanNumber) {
    return {
      id: createTemporaryUploadId(),
      file,
      fileName: prefixedName,
      pathKey: "",
      folderPath: "",
      sponsorId: loanNumber
        ? folders.find((folder) => folder.name === loanNumber)?.sponsorId ??
          null
        : null,
      tagCategory: null,
      tags: [],
      pathVerified: false,
    };
  }

  const destination = resolveSamePathDestination(
    loanNumber,
    structureRelativePath,
    folders,
  );

  if (!destination) {
    return {
      id: createTemporaryUploadId(),
      file,
      fileName: prefixedName,
      pathKey: "",
      folderPath: "",
      sponsorId: null,
      tagCategory: null,
      tags: [],
      pathVerified: false,
    };
  }

  const tagCategory = determineTagCategoryFromFolderPath(destination.folderPath);
  const automaticTag = tagCategory
    ? automaticFolderTags[tagCategory]
    : undefined;

  return {
    id: createTemporaryUploadId(),
    file,
    fileName: prefixedName,
    pathKey: destination.pathKey,
    folderPath: destination.folderPath,
    sponsorId: destination.sponsorId,
    tagCategory,
    tags: automaticTag ? [{ ...automaticTag }] : [],
    pathVerified: false,
  };
};

const remapSamePathUploadItems = (
  files: IUploadFileItem[],
  structureRelativePath: string,
  folders: ILoanFolder[],
  automaticFolderTags: Partial<Record<TaxonomyCategory, ITaxonomyTag>>,
): IUploadFileItem[] => {
  if (!structureRelativePath) {
    return files.map((item) => {
      const loanNumber = parseLoanNumberFromFileName(item.file.name, folders);
      return {
        ...item,
        fileName: loanNumber
          ? buildSamePathPrefixedFileName(loanNumber, item.file.name)
          : item.file.name,
        pathKey: "",
        folderPath: "",
        sponsorId: loanNumber
          ? folders.find((folder) => folder.name === loanNumber)?.sponsorId ??
            null
          : null,
        tagCategory: null,
        tags: [],
        pathVerified: false,
      };
    });
  }

  return files.map((item) => {
    const loanNumber = parseLoanNumberFromFileName(item.file.name, folders);
    const prefixedName = loanNumber
      ? buildSamePathPrefixedFileName(loanNumber, item.file.name)
      : item.file.name;

    if (!loanNumber) {
      return {
        ...item,
        fileName: prefixedName,
        pathKey: "",
        folderPath: "",
        sponsorId: null,
        tagCategory: null,
        tags: [],
        pathVerified: false,
      };
    }

    const destination = resolveSamePathDestination(
      loanNumber,
      structureRelativePath,
      folders,
    );

    if (!destination) {
      return {
        ...item,
        fileName: prefixedName,
        pathKey: "",
        folderPath: "",
        sponsorId: null,
        tagCategory: null,
        tags: [],
        pathVerified: false,
      };
    }

    const tagCategory = determineTagCategoryFromFolderPath(destination.folderPath);
    const automaticTag = tagCategory
      ? automaticFolderTags[tagCategory]
      : undefined;

    return {
      ...item,
      fileName: prefixedName,
      pathKey: destination.pathKey,
      folderPath: destination.folderPath,
      sponsorId: destination.sponsorId,
      tagCategory,
      tags: automaticTag ? [{ ...automaticTag }] : [],
      pathVerified: false,
    };
  });
};

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
  const [bulkTagPickerResetKey, setBulkTagPickerResetKey] = useState(0);
  const [isDragOver, setIsDragOver] = useState(false);
  const [pathErrorFileIds, setPathErrorFileIds] = useState<string[]>([]);
  const [duplicatePathFileIds, setDuplicatePathFileIds] = useState<string[]>(
    [],
  );
  const [invalidFolderPathFileIds, setInvalidFolderPathFileIds] = useState<
    string[]
  >([]);
  const [uploadPathMode, setUploadPathMode] =
    useState<BulkUploadPathMode | null>(null);
  const [samePathFolderKey, setSamePathFolderKey] = useState("");
  const [samePathExpandedKeys, setSamePathExpandedKeys] = useState<
    Record<string, boolean>
  >({});
  const fileInputRef = useRef<HTMLInputElement>(null);
  const isMountedRef = useRef(true);
  const tagFieldsRef = useRef(tagFields);
  const tagTreesRef = useRef(tagTrees);
  const fileGridScrollRef = useRef<HTMLDivElement>(null);

  tagFieldsRef.current = tagFields;
  tagTreesRef.current = tagTrees;

  const pathDropdownOptions = useMemo(
    () => pathOptions.map((o) => ({ label: o.label, value: o.key })),
    [pathOptions],
  );

  const folderStructureTree = useMemo(
    () => buildFolderStructureTreeNodes(folderStructure as IFolderStructureNode[]),
    [],
  );

  const samePathUploadPreview = useMemo(() => {
    if (!samePathFolderKey) return "";
    return `*/${samePathFolderKey}`;
  }, [samePathFolderKey]);

  const samePathSelectionKeys = useMemo<string | null>(
    () => samePathFolderKey || null,
    [samePathFolderKey],
  );

  const renderSamePathFolderNode = useCallback(
    (node: TreeNode, options: { expanded: boolean }) => (
      <SamePathFolderNodeTemplate
        node={node}
        expanded={options.expanded}
        isSelected={samePathFolderKey === String(node.key)}
      />
    ),
    [samePathFolderKey],
  );

  const handleSamePathToggle = useCallback(
    (event: { value: Record<string, boolean> }) => {
      setSamePathExpandedKeys(event.value);
    },
    [],
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
              <FieldRef Name="${loanLibraryFields.sponsorName}" />
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
          sponsorId: parseLoanSponsorLookup(item).id || null,
        }))
        .filter(
          (folder: ILoanFolder) =>
            !!folder.name &&
            !!folder.fileRef &&
            !shouldSkipFolderName(folder.name),
        );

      const rootOptions: IPathOption[] = folders.map((folder) => ({
        key: folder.fileRef,
        label: folder.name,
        loanNumber: folder.name,
        folderPath: folder.fileRef,
        sponsorId: folder.sponsorId ?? null,
      }));
      let options: IPathOption[] = rootOptions;

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
      sponsorId: option?.sponsorId ?? null,
      tagCategory,
      tags: automaticTag ? [{ ...automaticTag }] : [],
      pathVerified: false,
    };
  };

  // Files
  // Add files with count limits and path auto-suggestions.
  const addFilesToUploadQueue = (files: File[]): void => {
    if (!files.length) return;

    if (uploadPathMode === "same") {
      if (!loanFolders.length) {
        toastFunc(
          "warn",
          "Warning",
          "Loan folders are still loading. Please wait and try again.",
        );
        return;
      }
    } else if (!pathOptions.length) {
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

    const accepted: IUploadFileItem[] = [];
    const duplicateNames: string[] = [];

    toAdd.forEach((file) => {
      const item =
        uploadPathMode === "same"
          ? createSamePathUploadQueueItem(
              file,
              samePathFolderKey,
              loanFolders,
              automaticFolderTags,
            )
          : createUploadQueueItem(file);

      const conflictsExisting = uploadFiles.some((existing) =>
        isSameFileSamePath(existing, item),
      );
      const conflictsAccepted = accepted.some((existing) =>
        isSameFileSamePath(existing, item),
      );

      if (item.folderPath && (conflictsExisting || conflictsAccepted)) {
        duplicateNames.push(file.name);
        return;
      }

      accepted.push(item);
    });

    if (duplicateNames.length) {
      const preview = Array.from(new Set(duplicateNames))
        .slice(0, 3)
        .map((name) => `"${name}"`)
        .join(", ");
      toastFunc(
        "warn",
        "Duplicate file path",
        duplicateNames.length === 1
          ? `${preview} is already mapped to the same folder path.`
          : `${duplicateNames.length} file(s) skipped — already mapped to the same folder path: ${preview}.`,
      );
    }

    if (!accepted.length) return;

    setUploadFiles((prev) => [...prev, ...accepted]);
    setDuplicatePathFileIds([]);
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
    const current = uploadFiles.find((item) => item.id === id);
    const nextPathKey = option?.key || folderPath;

    if (
      current &&
      folderPath &&
      uploadFiles.some(
        (item) =>
          item.id !== id &&
          isSameFileSamePath(item, {
            fileName: current.fileName,
            folderPath,
            pathKey: nextPathKey,
          }),
      )
    ) {
      const conflictingIds = uploadFiles
        .filter(
          (item) =>
            item.id === id ||
            isSameFileSamePath(item, {
              fileName: current.fileName,
              folderPath,
              pathKey: nextPathKey,
            }),
        )
        .map((item) => item.id);
      setDuplicatePathFileIds(conflictingIds);
      toastFunc(
        "warn",
        "Duplicate file path",
        `"${current.fileName}" is already mapped to this folder path. Choose a different path.`,
      );
      return;
    }

    const tagCategory = determineTagCategoryFromFolderPath(folderPath);
    const automaticTag = tagCategory
      ? automaticFolderTags[tagCategory]
      : undefined;
    const next = uploadFiles.map((item) =>
      item.id === id
        ? {
            ...item,
            pathKey: nextPathKey,
            folderPath,
            sponsorId: option?.sponsorId ?? null,
            tagCategory,
            tags: automaticTag ? [{ ...automaticTag }] : [],
            pathVerified: false,
          }
        : item,
    );
    setUploadFiles(next);
    setDuplicatePathFileIds(getDuplicateSamePathFileIds(next));
    if (folderPath) {
      setPathErrorFileIds((prev) => prev.filter((fileId) => fileId !== id));
    }
  };

  // Tags
  // Update the selected taxonomy category for one file (click again to clear).
  const updateUploadTagCategory = (
    id: string,
    category: TaxonomyCategory,
  ): void => {
    setUploadFiles((prev) =>
      prev.map((item) => {
        if (item.id !== id) return item;
        if (item.tagCategory === category) {
          return {
            ...item,
            tagCategory: null,
            tags: [],
          };
        }
        return {
          ...item,
          tagCategory: category,
          tags: [],
        };
      }),
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
    const next = uploadFiles.filter((item) => item.id !== id);
    setUploadFiles(next);
    setPathErrorFileIds((prev) => prev.filter((fileId) => fileId !== id));
    setDuplicatePathFileIds(getDuplicateSamePathFileIds(next));
  };

  // Reset wizard state and clear all upload files.
  const clearUploadQueue = (): void => {
    setUploadFiles([]);
    setPathErrorFileIds([]);
    setDuplicatePathFileIds([]);
    setInvalidFolderPathFileIds([]);
    setBulkTagCategory(null);
    setBulkTagTermIds([]);
  };

  const resetWizardToStart = (): void => {
    clearUploadQueue();
    setUploadPathMode(null);
    setSamePathFolderKey("");
    setSamePathExpandedKeys({});
    setActiveStep(1);
  };

  const handleUploadPathModeChange = (mode: BulkUploadPathMode): void => {
    if (uploadPathMode && mode !== uploadPathMode && uploadFiles.length) {
      clearUploadQueue();
    }
    setUploadPathMode(mode);
    if (mode === "different") {
      setSamePathFolderKey("");
      setSamePathExpandedKeys({});
    }
  };

  const handleSamePathFolderSelect = useCallback(
    (structureRelativePath: string): void => {
      if (!structureRelativePath) return;

      setSamePathFolderKey(structureRelativePath);
      setInvalidFolderPathFileIds([]);

      if (!uploadFiles.length) {
        return;
      }

      const next = remapSamePathUploadItems(
        uploadFiles,
        structureRelativePath,
        loanFolders,
        automaticFolderTags,
      );
      setUploadFiles(next);
      setPathErrorFileIds([]);
      setDuplicatePathFileIds(getDuplicateSamePathFileIds(next));
    },
    [automaticFolderTags, loanFolders, uploadFiles],
  );

  const resolveSamePathSelectionKey = (
    value: TreeSelectionEvent["value"],
  ): string | undefined => {
    if (!value) {
      return undefined;
    }

    if (typeof value === "string") {
      return value;
    }

    if (typeof value === "object") {
      return Object.keys(value).find(
        (key) => !!(value as Record<string, boolean>)[key],
      );
    }

    return undefined;
  };

  const handleSamePathSelectionChange = useCallback(
    (event: TreeSelectionEvent) => {
      const selectedKey = resolveSamePathSelectionKey(event.value);

      if (!selectedKey) {
        return;
      }

      handleSamePathFolderSelect(selectedKey);

      const node = findFolderStructureNode(folderStructureTree, selectedKey);
      if (node?.children?.length) {
        setSamePathExpandedKeys((previousKeys) =>
          previousKeys[selectedKey]
            ? previousKeys
            : { ...previousKeys, [selectedKey]: true },
        );
      }
    },
    [folderStructureTree, handleSamePathFolderSelect],
  );

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

    // Reset bulk apply bar so category pills + dropdown clear after apply.
    setBulkTagCategory(null);
    setBulkTagTermIds([]);
    setBulkTagPickerResetKey((k) => k + 1);

    toastFunc(
      "success",
      "Success",
      `${tags.length} tag(s) applied to ${uploadFiles.length} file(s).`,
    );
  };

  // Resolve display label for selected folder path.
  const getFolderPathDisplayLabel = (folderPath: string): string => {
    if (uploadPathMode === "same") {
      return formatSamePathFromFolderPath(folderPath, loanFolders);
    }

    return (
      pathOptions.find((option) =>
        areSharePointPathsEqual(option.key, folderPath),
      )?.label || formatFolderPathLabel(folderPath, loanFolders)
    );
  };

  // Verify that a target folder path exists in SharePoint.
  const doesFolderPathExist = async (folderPath: string): Promise<boolean> => {
    const serverRelativePath = ensureServerRelativePath(folderPath);
    if (!serverRelativePath) return false;

    try {
      const folder: any = await sp.web
        .getFolderByServerRelativePath(serverRelativePath)
        .select("Exists", "ServerRelativeUrl")();
      return (
        folder?.Exists !== false &&
        !!String(folder?.ServerRelativeUrl || "").trim()
      );
    } catch {
      return false;
    }
  };

  // Verify all selected folder paths before tags/submit.
  const verifyAllFolderPaths = async (
    files: IUploadFileItem[] = uploadFiles,
  ): Promise<{
    allValid: boolean;
    failedPaths: string[];
  }> => {
    const uniquePaths = Array.from(
      new Set(
        files
          .map((item) => ensureServerRelativePath(item.folderPath))
          .filter((path) => !!path),
      ),
    );

    if (!uniquePaths.length) {
      return { allValid: false, failedPaths: [] };
    }

    const pathResults = await Promise.all(
      uniquePaths.map(async (path) => ({
        path,
        verified: await doesFolderPathExist(path),
      })),
    );

    const verifiedMap = new Map(
      pathResults.map((result) => [result.path, result.verified]),
    );
    const failedPaths = pathResults
      .filter((result) => !result.verified)
      .map((result) => result.path);

    setUploadFiles((prev) =>
      prev.map((item) => {
        const normalizedPath = ensureServerRelativePath(item.folderPath);
        return {
          ...item,
          pathVerified: !!verifiedMap.get(normalizedPath),
        };
      }),
    );

    setInvalidFolderPathFileIds(
      files
        .filter((item) => {
          const normalizedPath = ensureServerRelativePath(item.folderPath);
          return normalizedPath && failedPaths.includes(normalizedPath);
        })
        .map((item) => item.id),
    );

    return {
      allValid: failedPaths.length === 0,
      failedPaths,
    };
  };

  const showFolderDoesNotExistMessage = (failedPaths: string[]): void => {
    const labels = Array.from(
      new Set(
        failedPaths.map((folderPath) => getFolderPathDisplayLabel(folderPath)),
      ),
    ).slice(0, 5);

    if (!labels.length) {
      toastFunc(
        "error",
        "Folder does not exist",
        "This path is not in the SharePoint library. Please create the folder and try again.",
      );
      return;
    }

    toastFunc(
      "error",
      "Folder does not exist",
      labels.length === 1
        ? `Folder does not exist: ${labels[0]}. This path is not in the SharePoint library. Please create the folder and try again.`
        : `The following folders do not exist in the SharePoint library: ${labels.join(", ")}. Please create the folders and try again.`,
    );
  };

  const verifyFolderPathsFirst = async (
    files: IUploadFileItem[] = uploadFiles,
  ): Promise<boolean> => {
    const hasMappedPaths = files.some((item) => !!item.folderPath);
    if (!hasMappedPaths) {
      toastFunc(
        "warn",
        "Path required",
        "Select a folder path before continuing.",
      );
      return false;
    }

    const { allValid, failedPaths } = await verifyAllFolderPaths(files);
    if (!allValid) {
      showFolderDoesNotExistMessage(failedPaths);
      return false;
    }

    setInvalidFolderPathFileIds([]);
    return true;
  };

  // Validate
  // Validate path mode selection before upload step.
  const validateUploadModeStep = (): boolean => {
    if (!uploadPathMode) {
      toastFunc("warn", "Warning", "Select an upload path type to continue");
      return false;
    }
    return true;
  };

  // Validate
  // Validate step one inputs before opening tags step.
  const validateFileMappingStep = (): boolean => {
    if (!uploadFiles.length) {
      toastFunc("warn", "Warning", "Add at least one file to continue");
      return false;
    }

    if (uploadPathMode === "same") {
      if (!samePathFolderKey) {
        toastFunc(
          "warn",
          "Folder path required",
          "Select a folder from the folder structure (e.g. Asset Management → Underwriting).",
        );
        return false;
      }

      const missingLoanFiles = uploadFiles.filter(
        (file) => !parseLoanNumberFromFileName(file.file.name, loanFolders),
      );

      if (missingLoanFiles.length) {
        const missingIds = missingLoanFiles.map((file) => file.id);
        setPathErrorFileIds(missingIds);
        setDuplicatePathFileIds([]);

        toastFunc(
          "warn",
          "Loan number required",
          missingLoanFiles.length === 1
            ? `"${missingLoanFiles[0].file.name}" must start with a valid loan number followed by an underscore (e.g. 3000101_Appraisal.pdf).`
            : `${missingLoanFiles.length} files are missing a valid loan number prefix. Use format: {loanNumber}_{fileName}.`,
        );

        window.setTimeout(() => {
          const firstErrorCard = fileGridScrollRef.current?.querySelector(
            `[data-file-id="${missingIds[0]}"]`,
          ) as HTMLElement | null;
          firstErrorCard?.scrollIntoView({ behavior: "smooth", block: "nearest" });
        }, 50);

        return false;
      }
    }

    const missingPathFiles = uploadFiles.filter(
      (file) => !file.pathKey || !file.folderPath,
    );

    if (missingPathFiles.length) {
      const missingIds = missingPathFiles.map((file) => file.id);
      setPathErrorFileIds(missingIds);
      setDuplicatePathFileIds([]);

      const previewNames = missingPathFiles
        .slice(0, 3)
        .map((file) => `"${file.fileName}"`)
        .join(", ");
      const moreCount = missingPathFiles.length - 3;

      if (uploadPathMode === "same") {
        const pathExamples = missingPathFiles
          .slice(0, 3)
          .map((file) => {
            const loanNumber = parseLoanNumberFromFileName(
              file.file.name,
              loanFolders,
            );
            return loanNumber && samePathFolderKey
              ? formatSamePathDisplayLabel(loanNumber, samePathFolderKey)
              : `"${file.fileName}"`;
          })
          .join(", ");

        toastFunc(
          "warn",
          "Folder path not found",
          missingPathFiles.length === 1
            ? `Could not resolve upload path ${pathExamples}. Verify the loan folder exists in SharePoint.`
            : `${missingPathFiles.length} files could not be mapped (${pathExamples}${moreCount > 0 ? `, and ${moreCount} more` : ""}). Verify loan folders exist in SharePoint.`,
        );
      } else {
        toastFunc(
          "warn",
          "Path required",
          moreCount > 0
            ? `${missingPathFiles.length} files need a folder path. Missing: ${previewNames}, and ${moreCount} more.`
            : missingPathFiles.length === 1
              ? `${previewNames} needs a folder path. Select Path / Location to continue.`
              : `${missingPathFiles.length} files need a folder path: ${previewNames}.`,
        );
      }

      window.setTimeout(() => {
        const firstErrorCard = fileGridScrollRef.current?.querySelector(
          `[data-file-id="${missingIds[0]}"]`,
        ) as HTMLElement | null;
        firstErrorCard?.scrollIntoView({ behavior: "smooth", block: "nearest" });
      }, 50);

      return false;
    }

    const duplicateIds = getDuplicateSamePathFileIds(uploadFiles);
    if (duplicateIds.length) {
      setPathErrorFileIds([]);
      setDuplicatePathFileIds(duplicateIds);

      const duplicateFiles = uploadFiles.filter((file) =>
        duplicateIds.includes(file.id),
      );
      const uniqueNames = Array.from(
        new Set(duplicateFiles.map((file) => file.fileName)),
      );
      const previewNames = uniqueNames
        .slice(0, 3)
        .map((name) => `"${name}"`)
        .join(", ");

      toastFunc(
        "warn",
        "Duplicate file path",
        uniqueNames.length === 1
          ? `${previewNames} is mapped to the same folder more than once. Change one path or remove the duplicate.`
          : `Same file name cannot be mapped to the same folder twice: ${previewNames}.`,
      );

      window.setTimeout(() => {
        const firstErrorCard = fileGridScrollRef.current?.querySelector(
          `[data-file-id="${duplicateIds[0]}"]`,
        ) as HTMLElement | null;
        firstErrorCard?.scrollIntoView({ behavior: "smooth", block: "nearest" });
      }, 50);

      return false;
    }

    setPathErrorFileIds([]);
    setDuplicatePathFileIds([]);
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

  // Validate all requirements before showing the upload progress popup.
  const validateBeforeSubmit = async (): Promise<boolean> => {
    if (!uploadFiles.length) {
      toastFunc("warn", "No files to upload", "Add at least one file before submitting.");
      return false;
    }

    if (!validateFileMappingStep()) {
      return false;
    }

    // Path existence check runs first before tags and upload popup.
    if (!(await verifyFolderPathsFirst())) {
      return false;
    }

    if (!validateTagAssignmentStep()) {
      return false;
    }

    return true;
  };

  // Steps
  // Move from path mode step to upload step.
  const navigateFromModeStep = (): void => {
    if (!validateUploadModeStep()) return;
    setActiveStep(2);
  };

  // Steps
  // Move from mapping step to tag assignment step.
  const navigateToTagAssignmentStep = async (): Promise<void> => {
    if (!validateFileMappingStep() || verifyingPaths) return;

    setVerifyingPaths(true);
    try {
      if (!(await verifyFolderPathsFirst())) {
        return;
      }

      setActiveStep(3);
    } finally {
      setVerifyingPaths(false);
    }
  };

  // Steps
  // Move from tag step to review step after validation.
  const navigateToReviewStep = async (): Promise<void> => {
    if (!validateTagAssignmentStep() || verifyingPaths) return;

    setVerifyingPaths(true);
    try {
      if (!(await verifyFolderPathsFirst())) {
        return;
      }

      setActiveStep(4);
    } finally {
      setVerifyingPaths(false);
    }
  };

  // Submit
  // Run all validations first; only show upload popup when everything passes.
  const handleSubmit = async (): Promise<void> => {
    if (!uploadFiles.length || submitting || verifyingPaths) return;

    setVerifyingPaths(true);
    try {
      const isValid = await validateBeforeSubmit();
      if (!isValid) return;
    } finally {
      setVerifyingPaths(false);
    }

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
        const failedResults = results.filter((result) => !result.success);
        const folderMissingResults = failedResults.filter(
          (result) =>
            result.folderNotFound ||
            isSharePointFolderNotFoundError(result.error),
        );

        if (folderMissingResults.length > 0) {
          const failedPaths = Array.from(
            new Set(
              uploadFiles
                .filter((item) =>
                  failedResults.some(
                    (result) => result.fileName === item.fileName,
                  ),
                )
                .map((item) => ensureServerRelativePath(item.folderPath))
                .filter((path) => !!path),
            ),
          );

          setInvalidFolderPathFileIds(
            uploadFiles
              .filter((item) =>
                failedResults.some(
                  (result) => result.fileName === item.fileName,
                ),
              )
              .map((item) => item.id),
          );

          showFolderDoesNotExistMessage(failedPaths);
        } else {
          const failedNames = failedResults
            .slice(0, 3)
            .map((result) => result.fileName)
            .join(", ");
          const suffix =
            failCount > 3 ? ` and ${failCount - 3} more` : "";

          toastFunc(
            "error",
            "Upload failed",
            `${failCount} file(s) failed to upload${failedNames ? `: ${failedNames}${suffix}` : ""}`,
          );
        }
      }

      if (successCount > 0 && failCount === 0) {
        await new Promise((resolve) => setTimeout(resolve, 800));
        resetWizardToStart();
      }
    } finally {
      setSubmitting(false);
      setUploadProgress(null);
    }
  };

  // Cancel the workflow and reset wizard to first step.
  const handleCancel = (): void => {
    resetWizardToStart();
  };

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

  if (loader && !uploadFiles.length && activeStep === 2) {
    return <Loader />;
  }

  const renderDifferentPathUploadStep = (): React.ReactElement => (
    <>
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
        <div className={styles.fileGridScroll} ref={fileGridScrollRef}>
          {pathErrorFileIds.length > 0 && (
            <div className={styles.pathValidationBanner} role="alert">
              <i className="pi pi-exclamation-circle" />
              <span>
                {pathErrorFileIds.length === 1
                  ? "1 file is missing a folder path. Highlighted below."
                  : `${pathErrorFileIds.length} files are missing a folder path. Highlighted below.`}
              </span>
            </div>
          )}
          {duplicatePathFileIds.length > 0 && (
            <div className={styles.pathValidationBanner} role="alert">
              <i className="pi pi-exclamation-circle" />
              <span>
                Same file name cannot be mapped to the same folder more than
                once. Highlighted below.
              </span>
            </div>
          )}
          <div className={styles.fileGrid}>
          {uploadFiles.map((item) => {
            const hasPathError = pathErrorFileIds.includes(item.id);
            const hasDuplicatePathError = duplicatePathFileIds.includes(
              item.id,
            );
            const hasMappingError = hasPathError || hasDuplicatePathError;
            return (
            <div
              key={item.id}
              data-file-id={item.id}
              className={`${styles.fileCard} ${
                hasMappingError ? styles.fileCardPathError : ""
              }`}
            >
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
                {hasPathError && (
                  <span className={styles.pathErrorBadge}>Path needed</span>
                )}
                {hasDuplicatePathError && (
                  <span className={styles.pathErrorBadge}>Duplicate</span>
                )}
              </div>
              <label
                className={`${styles.fieldLabel} ${
                  hasMappingError ? styles.fieldLabelError : ""
                }`}
              >
                Path / Location {hasMappingError ? "*" : ""}
              </label>
              <UploadPathDropdown
                value={item.pathKey}
                options={pathDropdownOptions}
                hasError={hasMappingError}
                disabled={!pathDropdownOptions.length}
                onChange={(pathKey) => updateUploadDestination(item.id, pathKey)}
              />
              {hasPathError && (
                <div className={styles.pathErrorMessage}>
                  <i className="pi pi-info-circle" />
                  Please select a folder path for this file
                </div>
              )}
              {hasDuplicatePathError && (
                <div className={styles.pathErrorMessage}>
                  <i className="pi pi-info-circle" />
                  Same file is already mapped to this folder — change path
                  or remove duplicate
                </div>
              )}
            </div>
            );
          })}
          </div>
        </div>
      )}

      <div className={styles.stepFooter}>
        <Button
          label="← Back"
          outlined
          className={styles.backBtn}
          onClick={() => setActiveStep(1)}
        />
        {uploadFiles.length > 0 && (
          <Button
            label="Clear all"
            outlined
            className={styles.clearBtn}
            onClick={clearUploadQueue}
          />
        )}
        <Button
          label={verifyingPaths ? "Verifying folders..." : "Next: Assign Tags →"}
          className={styles.primaryBtn}
          icon={verifyingPaths ? "pi pi-spin pi-spinner" : undefined}
          onClick={() => void navigateToTagAssignmentStep()}
          disabled={!uploadFiles.length || verifyingPaths}
        />
      </div>
    </>
  );

  const renderSamePathUploadStep = (): React.ReactElement => (
    <>
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
          Supports up to {MAX_FILES} files — file names must use{" "}
          {"{loanNumber}_{fileName}"} (e.g. 3031101_Appraisal.pdf)
        </div>
        <input
          ref={fileInputRef}
          type="file"
          multiple
          className={styles.hiddenInput}
          onChange={handleFileSelect}
        />
      </div>

      <div className={styles.samePathLayout}>
        <div className={styles.samePathPanel}>
          <p className={styles.copyMoveSectionLabel}>Select Folder Path</p>
          <p className={styles.copyMoveHint}>
            One folder path applies to all files. Each file&apos;s loan number
            is read from its file name.
          </p>
          {samePathUploadPreview && (
            <div className={styles.samePathPreview}>
              <div>
                <strong>Upload Path:</strong> {samePathUploadPreview}
              </div>
              <div>
                <strong>File Name:</strong> {"{loanNumber}_<original-file-name>"}
              </div>
            </div>
          )}
          <div className={styles.copyMoveTreePanel}>
            <Tree
              value={folderStructureTree}
              expandedKeys={samePathExpandedKeys}
              onToggle={handleSamePathToggle}
              selectionMode="single"
              selectionKeys={samePathSelectionKeys}
              onSelectionChange={handleSamePathSelectionChange}
              metaKeySelection={false}
              nodeTemplate={renderSamePathFolderNode}
              className={styles.copyMoveTree}
            />
          </div>
        </div>

        <div className={styles.samePathPanel}>
          {uploadFiles.length > 0 ? (
            <div className={styles.fileGridScroll} ref={fileGridScrollRef}>
              {!samePathFolderKey && (
                <div className={styles.pathValidationBanner} role="alert">
                  <i className="pi pi-exclamation-circle" />
                  <span>
                    Select a folder path from the folder structure (e.g. Asset
                    Management → Underwriting).
                  </span>
                </div>
              )}
              {uploadFiles.some(
                (file) =>
                  getSamePathFileIssue(
                    file,
                    loanFolders,
                    samePathFolderKey,
                    invalidFolderPathFileIds,
                  ) === "missing_loan_number",
              ) && (
                <div className={styles.pathValidationBanner} role="alert">
                  <i className="pi pi-exclamation-circle" />
                  <span>
                    One or more files are missing a valid loan number prefix.
                    Use format: {"{loanNumber}_{fileName}"} (e.g.
                    3000101_Appraisal.pdf).
                  </span>
                </div>
              )}
              {uploadFiles.some(
                (file) =>
                  getSamePathFileIssue(
                    file,
                    loanFolders,
                    samePathFolderKey,
                    invalidFolderPathFileIds,
                  ) === "unresolved_path",
              ) && (
                <div className={styles.pathValidationBanner} role="alert">
                  <i className="pi pi-exclamation-circle" />
                  <span>
                    One or more folder paths could not be resolved. Verify the
                    loan folder exists in SharePoint.
                  </span>
                </div>
              )}
              {uploadFiles.some(
                (file) =>
                  getSamePathFileIssue(
                    file,
                    loanFolders,
                    samePathFolderKey,
                    invalidFolderPathFileIds,
                  ) === "folder_not_in_sharepoint",
              ) && (
                <div className={styles.pathValidationBanner} role="alert">
                  <i className="pi pi-exclamation-circle" />
                  <span>
                    The selected path does not exist in SharePoint. Please create
                    the folder and try again.
                  </span>
                </div>
              )}
              {duplicatePathFileIds.length > 0 && (
                <div className={styles.pathValidationBanner} role="alert">
                  <i className="pi pi-exclamation-circle" />
                  <span>
                    Same file name cannot be uploaded to the same folder more
                    than once. Highlighted below.
                  </span>
                </div>
              )}
              <div className={styles.fileGrid}>
                {uploadFiles.map((item) => {
                  const loanNumber = parseLoanNumberFromFileName(
                    item.file.name,
                    loanFolders,
                  );
                  const fileIssue = getSamePathFileIssue(
                    item,
                    loanFolders,
                    samePathFolderKey,
                    invalidFolderPathFileIds,
                  );
                  const hasDuplicatePathError = duplicatePathFileIds.includes(
                    item.id,
                  );
                  const hasMappingError = !!fileIssue || hasDuplicatePathError;
                  const resolvedPath = item.folderPath
                    ? formatSamePathFromFolderPath(item.folderPath, loanFolders)
                    : loanNumber && samePathFolderKey
                      ? formatSamePathDisplayLabel(
                          loanNumber,
                          samePathFolderKey,
                        )
                      : loanNumber
                        ? "Select a folder path"
                        : "Loan number not found in file name";
                  const issueMessage = fileIssue
                    ? getSamePathFileIssueMessage(
                        fileIssue,
                        loanNumber,
                        samePathFolderKey,
                        item.folderPath,
                        loanFolders,
                      )
                    : "";

                  return (
                    <div
                      key={item.id}
                      data-file-id={item.id}
                      className={`${styles.fileCard} ${
                        hasMappingError ? styles.fileCardPathError : ""
                      }`}
                    >
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
                        {fileIssue === "missing_loan_number" && (
                          <span className={styles.pathErrorBadge}>
                            Loan needed
                          </span>
                        )}
                        {fileIssue === "missing_folder_selection" && (
                          <span className={styles.pathErrorBadge}>
                            Path needed
                          </span>
                        )}
                        {fileIssue === "unresolved_path" && (
                          <span className={styles.pathErrorBadge}>
                            Path not found
                          </span>
                        )}
                        {fileIssue === "folder_not_in_sharepoint" && (
                          <span className={styles.pathErrorBadge}>
                            Folder missing
                          </span>
                        )}
                        {hasDuplicatePathError && (
                          <span className={styles.pathErrorBadge}>Duplicate</span>
                        )}
                      </div>
                      <div className={styles.reviewPath} title={resolvedPath}>
                        {resolvedPath}
                      </div>
                      {issueMessage && (
                        <div className={styles.pathErrorMessage}>
                          <i className="pi pi-info-circle" />
                          {issueMessage}
                        </div>
                      )}
                      {hasDuplicatePathError && (
                        <div className={styles.pathErrorMessage}>
                          <i className="pi pi-info-circle" />
                          Same file is already queued for this folder — remove
                          duplicate
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          ) : (
            <div className={styles.samePathEmptyFiles}>
              Uploaded files will appear here with their loan number and path.
            </div>
          )}
        </div>
      </div>

      <div className={styles.stepFooter}>
        <Button
          label="← Back"
          outlined
          className={styles.backBtn}
          onClick={() => setActiveStep(1)}
        />
        {uploadFiles.length > 0 && (
          <Button
            label="Clear all"
            outlined
            className={styles.clearBtn}
            onClick={clearUploadQueue}
          />
        )}
        <Button
          label={verifyingPaths ? "Verifying folders..." : "Next: Assign Tags →"}
          className={styles.primaryBtn}
          icon={verifyingPaths ? "pi pi-spin pi-spinner" : undefined}
          onClick={() => void navigateToTagAssignmentStep()}
          disabled={!uploadFiles.length || verifyingPaths}
        />
      </div>
    </>
  );

  return (
    <div className={styles.bulkUploadWrapper}>
      {renderSubmitProgress()}

      <div className={styles.stepper}>
        {renderStepIndicator(1, "Upload Path")}
        <div className={styles.connector} />
        {renderStepIndicator(2, "Select & Map Files")}
        <div className={styles.connector} />
        {renderStepIndicator(3, "Assign Tags")}
        <div className={styles.connector} />
        {renderStepIndicator(4, "Review & Submit")}
      </div>

      {/* ─── Step 1: Upload Path Mode ───────────────────────────────────── */}
      {activeStep === 1 && (
        <div className={styles.stepContent}>
          <div className={styles.pathModeSection}>
            <div className={styles.pathModeHeader}>
              <h3 className={styles.pathModeTitle}>How would you like to upload?</h3>
              <p className={styles.pathModeHint}>
                Choose the upload workflow that fits your files.
              </p>
            </div>

            <div className={styles.pathModeOptions} role="radiogroup" aria-label="Upload path type">
              {PATH_MODE_OPTIONS.map((option) => {
                const isSelected = uploadPathMode === option.value;
                return (
                  <button
                    key={option.value}
                    type="button"
                    role="radio"
                    aria-checked={isSelected}
                    className={`${styles.pathModeOption} ${
                      isSelected ? styles.pathModeOptionSelected : ""
                    }`}
                    onClick={() =>
                      handleUploadPathModeChange(option.value)
                    }
                  >
                    <span className={styles.pathModeOptionIconWrap}>
                      <i className={`${option.icon} ${styles.pathModeOptionIcon}`} />
                    </span>
                    <span className={styles.pathModeOptionBody}>
                      <span className={styles.pathModeOptionTitle}>
                        {option.title}
                      </span>
                      <span className={styles.pathModeOptionDesc}>
                        {option.description}
                      </span>
                      <span className={styles.pathModeOptionExample}>
                        {option.example}
                      </span>
                    </span>
                    <span
                      className={`${styles.pathModeOptionCheck} ${
                        isSelected ? styles.pathModeOptionCheckSelected : ""
                      }`}
                      aria-hidden="true"
                    >
                      {isSelected ? <i className="pi pi-check" /> : null}
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          <div className={styles.stepFooter}>
            <Button
              label="Next →"
              className={styles.primaryBtn}
              onClick={navigateFromModeStep}
              disabled={!uploadPathMode}
            />
          </div>
        </div>
      )}

      {/* ─── Step 2: Select & Map Files ─────────────────────────────────── */}
      {activeStep === 2 && (
        <div className={styles.stepContent}>
          {uploadPathMode === "same"
            ? renderSamePathUploadStep()
            : renderDifferentPathUploadStep()}
        </div>
      )}

      {/* ─── Step 3: Assign Tags ────────────────────────────────────────── */}
      {activeStep === 3 && (
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
                    // Click active category again to clear (no tag).
                    if (bulkTagCategory === tag) {
                      setBulkTagCategory(null);
                      setBulkTagTermIds([]);
                      setBulkTagPickerResetKey((k) => k + 1);
                      return;
                    }
                    setBulkTagCategory(tag);
                    setBulkTagTermIds([]);
                    setBulkTagPickerResetKey((k) => k + 1);
                  }}
                  title={
                    bulkTagCategory === tag
                      ? "Click again to clear category"
                      : `Select ${tag}`
                  }
                >
                  {tag}
                </button>
              ))}
            </div>
            <div className={styles.bulkTagPicker}>
              <TaxonomyTagPicker
                key={`bulk-${bulkTagCategory || "none"}-${bulkTagPickerResetKey}`}
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

          <div className={styles.fileGridScroll}>
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
                      title={
                        item.tagCategory === tag
                          ? "Click again to clear category"
                          : `Select ${tag}`
                      }
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
          </div>

          <div className={styles.stepFooter}>
            <Button
              label="← Back"
              outlined
              className={styles.backBtn}
              onClick={() => setActiveStep(2)}
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

      {/* ─── Step 4: Review & Submit ────────────────────────────────────── */}
      {activeStep === 4 && (
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
              onClick={() => setActiveStep(3)}
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
                  verifyingPaths
                    ? "Checking..."
                    : submitting
                      ? `Uploading ${uploadProgress?.completed || 0} of ${uploadFiles.length}...`
                      : `Submit ${uploadFiles.length} File${
                          uploadFiles.length !== 1 ? "s" : ""
                        }`
                }
                icon={
                  verifyingPaths || submitting
                    ? "pi pi-spin pi-spinner"
                    : "pi pi-check"
                }
                className={styles.primaryBtn}
                onClick={() => void handleSubmit()}
                disabled={submitting || verifyingPaths}
              />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default BulkUpload;
