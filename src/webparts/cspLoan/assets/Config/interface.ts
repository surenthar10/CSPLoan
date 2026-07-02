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
  createdby: string;
  date: Date | null;
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
  createddate: Date | null;
  modifieddate: Date | null;
  assetmanagement: string;
  servicing: string;
  legal: string;
  fileRef: string;
  fileName: string;
  folderType: number | null;
  serverRelativeUrl: string;
  type?: string;
}

export interface IMainData {
  isAdmin: boolean;
  loanDetails: ILoanRecord[];
}
