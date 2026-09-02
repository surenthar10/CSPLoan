import {
  IActions,
  IAllDropdowns,
  IBulkUploadConfig,
  IFlags,
  ISponsorUpdateConfig,
  IGroups,
  IListName,
  ILoanRecord,
  IPopup,
  ISponsorRecord,
} from "./interface";
import { Toast } from "primereact/toast";
import { ToastMessage } from "primereact/toast";
import { MutableRefObject } from "react";
export const toastRef: MutableRefObject<Toast | null> = { current: null };

export const flags: IFlags = {
  add: "add",
  edit: "edit",
  delete: "delete",
  view: "view",
  empty: "",
};
export const actions: IActions = {
  add: "Add",
  edit: "Edit",
  update: "Update",
  save: "Save",
  new: "New",
};
export const groupNames: IGroups = {
  adminGroup: "Admins",
  membersGroup: "Members",
};
export const listNames: IListName = {
  sponsors: "Sponsors",
  loan: "Test CSP Loan Files",
  loanFileMapping: "LoanFileMapping",
  loanFolderRequest: "LoanFolderRequests",
  loanActivityLog: "LoanActivityLog",
  // loan: "CSP Loan Files",
};

export const managedMetadataFields = {
  tags: "Tags",
  assetManagement: "Asset_x0020_management",
  legal: "Legal",
  servicing: "Servicing",
} as const;

export const loanLibraryFields = {
  sponsorName: "SponsorName",
  sponsorNameId: "SponsorNameId",
} as const;

export const parseLoanSponsorLookup = (
  item: Record<string, any> | null | undefined,
): { id: number; sponsorTitle: string } => {
  const lookup = item?.[loanLibraryFields.sponsorName];
  const entry = Array.isArray(lookup) ? lookup[0] : lookup;

  return {
    id: Number(
      entry?.lookupId || entry?.LookupId || entry?.Id || entry?.ID || 0,
    ),
    sponsorTitle: String(
      entry?.lookupValue || entry?.LookupValue || entry?.Title || "",
    ),
  };
};

export const buildSponsorLookupUpdatePayload = (
  sponsorId: number | null,
): Record<string, number | null> => ({
  [loanLibraryFields.sponsorNameId]: sponsorId,
});
export const bulkUploadConfig: IBulkUploadConfig = {
  maxFiles: 70,
  uploadConcurrency: 6,
};
export const sponsorUpdateConfig: ISponsorUpdateConfig = {
  batchSize: 100,
  batchConcurrency: 3,
};
export const confirmationPopup: IPopup = {
  Content: "",
  IsOpen: false,
  IsLoad: false,
};
// functions
export type ToastSeverity = "success" | "info" | "warn" | "error";

export const toastFunc = (
  severity: ToastSeverity,
  summary: string,
  detail: string,
): void => {
  if (toastRef.current) {
    const message: ToastMessage = {
      severity,
      summary,
      detail,
      life: 3000,
    };
    toastRef.current.show(message);
  }
};

export const errFunc = (err: any, funcName: string): void => {
  console.log(err, funcName);
  toastFunc("error", "Error", err);
};
export const deploymentConfig = (_siteUrl: string): string => {
  // const siteUrl: string = _siteUrl;
  let _CPSiteUrl: string = "";

  if (window.location.origin == "https://communityp.sharepoint.com/") {
    _CPSiteUrl = "https://communityp.sharepoint.com/sites/CSPTestLoantContent";
  }
  //else if (window.location.origin == "https://rocasanitario.sharepoint.com") {
  //   _CPSiteUrl = "https://rocasanitario.sharepoint.com/sites/RINMASTERDEV";
  // }
  return _CPSiteUrl;
};
export const fileSizeFinder = (fileSize: number): string => {
  let _size = fileSize;
  let fSExt = new Array("Bytes", "KB", "MB", "GB"),
    i = 0;
  while (_size > 900) {
    _size /= 1024;
    i++;
  }
  let exactSize = Math.round(_size * 100) / 100 + " " + fSExt[i];
  return exactSize;
};

export const sponsorColumns: ISponsorRecord = {
  id: null,
  sponsor: "",
  description: "",
  loans: [],
  createdby: "",
  date: null,
};
export const AllDropdowns: IAllDropdowns = {
  sponsor: [],
};
export const loanColumns: ILoanRecord = {
  id: null,
  name: "",
  sponsor: {
    id: null,
    sponsorTitle: "",
  },
  createdby: "",
  createddate: null,
  modifieddate: null,
  assetmanagement: "",
  servicing: "",
  legal: "",
  fileRef: "",
  fileName: "",
  folderType: null,
  serverRelativeUrl: "",
};
export const folderStructure: any = [
  {
    name: "Asset Management",
    children: [
      { name: "Borrower Contact Sheets" },
      { name: "Underwriting" },
      { name: "Financials and Rent Rolls" },
      { name: "Inspections + Enviro Reports" },
    ],
  },
  {
    name: "CSP Legal",
    children: [
      { name: "Pre Negotiation Agreements" },
      { name: "Servicing Mod" },
      {
        name: "FDIC Mod",
        children: [
          { name: "Closing Docs" },
          { name: "Closing Supporting Docs" },
          { name: "Business Diligence" },
        ],
      },
      { name: "Acceleration", children: [{ name: "Servicing Accel Package" }] },
      { name: "Foreclosure" },
    ],
  },
  { name: "Custodian" },
  { name: "Servicing", children: [{ name: "Insurance Compliance" }] },
  { name: "Notable Cases" },
];
