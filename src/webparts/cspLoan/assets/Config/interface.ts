export interface IPageDetails {
  name: string;
  data: any;
  path: string;
}

export interface IMainData {
  webUrl: string;
  tenantUrl: string;
  siteUrl: string;
  value: any[];
  pagedata: IPageDetails;
  rocaSiteUrl: string;
  confirmationPopup: IPopup;
  userRoles: string[];
}

export interface IGroups {
  adminGroup: string;
  membersGroup: string;
}

export interface IListName {
  sponsors: string;
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
  id: number;
  sponsor: string;
  description: string;
  loans: string[];
}
