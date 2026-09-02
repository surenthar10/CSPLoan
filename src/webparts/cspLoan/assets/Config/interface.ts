import { TreeNode } from "primereact/treenode";

export interface IPageDetails {
  name: string;
  data: any;
  path: string;
}

export interface IGroups {
  adminGroup: string;
  membersGroup: string;
}

export interface IListName {
  sponsors: string;
  loan: string;
  loanFileMapping: string;
  loanFolderRequest: string;
  loanActivityLog: string;
}

export interface IFlags {
  add: string;
  edit: string;
  view: string;
  delete: string;
  empty: string;
}

export interface IActions {
  add: string;
  edit: string;
  update: string;
  save: string;
  new: string;
}

export interface IDialogDetails {
  condition: boolean;
  data: any;
  type: string;
  inProcess?: boolean;
}

export interface IPopup {
  IsOpen: boolean;
  IsLoad: boolean;
  Content: string;
}

export interface IConfirmationPopupProps {
  IsOpen: boolean;
  Content: string;
  Loading: boolean;
  Life?: number;
  onClose: any;
  isMobileView: boolean;
}

export interface IFile {
  id: number | null;
  name: string;
  blob: any;
  isDelete: boolean;
  url: string;
  size?: number;
  docType?: string;
  type: string;
}

export interface ISponsorRecord {
  id: number | null;
  sponsor: string;
  description: string;
  loans: ILoanRecord[];
  loanSortLabel?: string;
  createdby: string;
  date: Date | string | null;
  modifieddate?: Date | string | null;
}
export interface IDrpdownOptions {
  code: number | string | null;
  name: string;
}
export interface IAllDropdowns {
  sponsor: IDrpdownOptions[];
}
export interface ILookuptype {
  id: number | null;
  sponsorTitle: string;
}
export interface ILoanRecord {
  id: number | null;
  name: string;
  sponsor: ILookuptype;
  createdby: string;
  createddate: string | null;
  modifieddate: string | null;
  assetmanagement: string;
  servicing: string;
  legal: string;
  fileRef: string;
  fileName: string;
  folderType: number | null;
  serverRelativeUrl: string;
  type?: string;
  isHyperlink?: boolean;
  mappingId?: number;
}

export interface ILoanFileMapping {
  Id: number;
  Title: string;
  SourceFileItemId: number;
  SourceFileUrl: string | { Url?: string; Description?: string };
  DestinationLoanItemId: number;
  RelativeFolderPath: string;
  IsActive?: boolean;
  Created?: Date | string;
  Modified?: Date | string;
  Author?: { Title?: string };
}

export interface IMainData {
  isAdmin: boolean;
  loanDetails: ILoanRecord[];
}

export type IBulkTaxonomyCategory =
  | "Asset Management"
  | "Legal"
  | "Servicing";

export interface IBulkUploadConfig {
  maxFiles: number;
  uploadConcurrency: number;
}

export interface ISponsorUpdateConfig {
  batchSize: number;
  batchConcurrency: number;
}

export interface ISponsorLoanGroup {
  loans: ILoanRecord[];
  sponsorId: number | null;
}

export interface IUploadProgress {
  total: number;
  completed: number;
  successCount: number;
  failCount: number;
  currentFileName: string;
  phase: "uploading" | "complete";
}

export interface IUploadResult {
  fileName: string;
  success: boolean;
  error?: string;
  folderNotFound?: boolean;
  folderPath?: string;
}

export interface ITaxonomyTag {
  termId: string;
  label: string;
  fieldInternalName: string;
  category: IBulkTaxonomyCategory;
}

export interface ITaxonomyFieldInfo {
  category: IBulkTaxonomyCategory;
  internalName: string;
  termSetId: string;
  groupId?: string;
}

export interface ITermStoreTerm {
  id?: string;
  Id?: string;
  labels?: { name: string; isDefault?: boolean }[];
  children?: ITermStoreTerm[];
  childrenCount?: number;
  parent?: { id?: string };
}

export interface ITaxonomyTagPickerProps {
  category: IBulkTaxonomyCategory | null;
  tree: TreeNode[];
  value: string[];
  displayLabels?: string[];
  placeholder?: string;
  disabled?: boolean;
  onChange: (termIds: string[]) => void;
}

export interface ILoanFolder {
  name: string;
  fileRef: string;
  sponsorId?: number | null;
}

export interface IPathOption {
  key: string;
  label: string;
  loanNumber: string;
  folderPath: string;
  sponsorId?: number | null;
}

export interface IUploadFileItem {
  id: string;
  file: File;
  fileName: string;
  pathKey: string;
  folderPath: string;
  sponsorId?: number | null;
  tagCategory: IBulkTaxonomyCategory | null;
  tags: ITaxonomyTag[];
  pathVerified: boolean;
}

export interface ILoanProps {
  context: any;
}

export interface ILoanFilterState {
  search: string;
  sponsor: IDrpdownOptions | null;
}

export interface IFolderDestination {
  Name: string;
  ServerRelativeUrl: string;
}

export interface IVersionHistoryRow {
  versionLabel: string;
  versionId?: number;
  isCurrent: boolean;
  name: string;
  sponsor: string;
  createdBy: string;
  createddate: string;
  modified: string;
  modifiedBy: string;
  assetmanagement: string;
  servicing: string;
  legal: string;
  size: string;
  comments: string;
  url?: string;
}

export type IVersionActionType = "restore" | "delete";

export interface IVersionActionDialog {
  visible: boolean;
  action: IVersionActionType | null;
  version: IVersionHistoryRow | null;
}
