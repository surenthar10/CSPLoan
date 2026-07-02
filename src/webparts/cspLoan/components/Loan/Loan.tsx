/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-use-before-define */

import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
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
  IDrpdownOptions,
  ILoanRecord,
} from "../../assets/Config/interface";
import { listNames, toastFunc } from "../../assets/Config/Config";
import styles from "./Loan.module.scss";
import { useDispatch } from "react-redux";
import { setLoanDetails } from "../../assets/Redux/Features/MainSPContextSlice";
import { FolderPicker } from "@pnp/spfx-controls-react/lib/FolderPicker";
// IFolder type from @pnp/spfx-controls-react may not be available in all setups.
// Define a minimal local IFolder interface to avoid module resolution errors.
interface IFolder {
  Name?: string;
  ServerRelativeUrl?: string;
  UniqueId?: string;
}

const AVATAR_COLOR = "#4f46e5";
// const LIBRARY_NAME = "exchange";
const LIBRARY_NAME = "testexchange";

// interface IAddLoan {
//   loannumber: string;
//   sponsor: { code: number | null; name: string };
// }

interface IProps {
  context: any;
}

interface IFilterState {
  search: string;
  sponsor: IDrpdownOptions | null;
}

const EMPTY_FILTER: IFilterState = { search: "", sponsor: null };

// ─── Shared CAML view fields used by both root + subfolder queries ─────────
const VIEW_FIELDS_XML = `
  <FieldRef Name="ID" /><FieldRef Name="FileRef" /><FieldRef Name="FileLeafRef" />
  <FieldRef Name="FSObjType" /><FieldRef Name="Author" /><FieldRef Name="Created" />
  <FieldRef Name="Modified" /><FieldRef Name="Sponsor" />
`;

// ─── Maps a raw renderListDataAsStream row into our ILoanRecord shape ──────
const mapRowToLoanRecord = (item: any): ILoanRecord => ({
  id: Number(item.ID),
  name: item.FileLeafRef || "",
  fileName: item.FileLeafRef || "",
  fileRef: item.FileRef || "",
  serverRelativeUrl: item.FileRef || "",
  folderType: Number(item.FSObjType) || 0,
  sponsor: {
    id: Number(item.Sponsor?.[0]?.lookupId || 0),
    sponsorTitle: item.Sponsor?.[0]?.lookupValue || "",
  },
  createdby: item.Author?.[0]?.title || item.Author?.title || "",
  createddate: item.Created || null,
  modifieddate: item.Modified || null,
  assetmanagement: "",
  servicing: "",
  legal: "",
  type: "existing",
});

const Loan = (props: IProps) => {
  // const initialLoanState: IAddLoan = {
  //   loannumber: "",
  //   sponsor: { code: null, name: "" },
  // };

  const dispatch: any = useDispatch();
  const [loader, setLoader] = useState(false);
  const [drpdown, setDrpdown] = useState<IAllDropdowns>({ sponsor: [] });
  const [mainFolders, setMainFolders] = useState<ILoanRecord[]>([]);

  // ─── displayItems = what's actually rendered in the table (post-filter) ────
  const [displayItems, setDisplayItems] = useState<ILoanRecord[]>([]);

  // ─── currentFolder = "" means we're at the Dashboard/root level ────────────
  const [currentFolder, setCurrentFolder] = useState("");

  // ─── filter state is PER LEVEL — reset whenever the level changes ──────────
  const [filter, setFilter] = useState<IFilterState>(EMPTY_FILTER);

  // ─── Folder contents cache — { folderPath -> fetched items } ──────────────
  // Populated on first visit to each folder. Lets Search/Filter/Reset and
  // breadcrumb back-navigation work instantly without re-hitting SharePoint.
  const folderCacheRef = useRef<Map<string, ILoanRecord[]>>(new Map());

  const menu = useRef<Menu>(null);
  const menuRef = useRef<Menu>(null);
  const selectedRowRef = useRef<any>(null);

  const [menuItems, setMenuItems] = useState<any[]>([]);
  const [selectedFolder, setSelectedFolder] = useState<any>(null);
  const [showRename, setShowRename] = useState(false);
  const [folderName, setFolderName] = useState("");
  const [showLinkDialog, setShowLinkDialog] = useState(false);
  const [showVersionHistory, setShowVersionHistory] = React.useState(false);
  const [versionHistory, setVersionHistory] = React.useState<any[]>([]);
  const [showCreateFolder, setShowCreateFolder] = useState(false);
  const [showDeleteDialog, setshowDeleteDialog] = useState(false);
  const [showCopyMoveDialog, setShowCopyMoveDialog] = useState(false);
  const [copyMoveOperation, setCopyMoveOperation] = useState<"copy" | "move">(
    "copy",
  );
  const [copyMoveTarget, setCopyMoveTarget] = useState<ILoanRecord | null>(
    null,
  );
  const [selectedDestination, setSelectedDestination] =
    useState<IFolder | null>(null);

  // ─── "New" dropdown menu items ─────────────────────────────────────────────
  const getNewMenuItem = (): void => {
    setMenuItems([
      // {
      //   label: "New Loan",
      //   icon: "pi pi-plus",
      //   visible: !currentFolder,
      //   command: () => setShowDialog(true),
      // },
      {
        label: "New sub folder",
        icon: "pi pi-folder-plus",
        visible: !!currentFolder,
        command: () => setShowCreateFolder(true),
      },
      {
        label: "File Upload",
        icon: "pi pi-file",
        command: () => console.log("File Upload clicked at:", currentFolder),
      },
    ]);
  };
  // ─── Row action menu items ──────────────────────────────────────────────────
  const getRowMenuItems = (rowData: any): void => {
    setMenuItems([
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
        icon: "pi pi-arrow-right-arrow-left",
        command: () => {
          openCopyMoveDialog(selectedRowRef.current, "move");
        },
      },
      {
        label: "Copy To",
        icon: "pi pi-copy",
        command: () => {
          openCopyMoveDialog(selectedRowRef.current, "copy");
        },
      },
      {
        label: "Rename",
        icon: "pi pi-pencil",
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
        command: () => {
          setshowDeleteDialog(true);
        },
      },
    ]);
  };
  // ─── Fetch ONLY root-level Loan Number folders (fast, no recursion) ───────
  const getLoanData = async (sponsorOptions: any[]): Promise<void> => {
    try {
      setLoader(true);

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
      folderCacheRef.current.clear();
    } catch (error) {
      console.error("getLoanData error:", error);
    } finally {
      setLoader(false);
    }
  };
  // ─── Fetch sponsor list (for filter dropdown) ──────────────────────────────
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
  // ─── Fetch a single folder's direct children (fast, scoped query) ─────────
  const fetchFolderContents = async (
    folderServerRelativeUrl: string,
  ): Promise<ILoanRecord[]> => {
    const list = sp.web.lists.getByTitle(listNames.loan);
    const response: any = await list.renderListDataAsStream({
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
    });

    return (response.Row || []).map(mapRowToLoanRecord);
  };
  // ─── Get the base (unfiltered) list for "" (root) or any folder — cache-first ──
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
  // ─── Apply search + sponsor filter against a base list ─────────────────────
  const applyFilter = (
    base: ILoanRecord[],
    state: IFilterState,
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
  // ═══════════════════════════════════════════════════════════════════════════
  // ─── SINGLE navigation function — Dashboard, any folder, any depth ───────
  // ═══════════════════════════════════════════════════════════════════════════
  const navigateToFolder = async (folderPath: string): Promise<void> => {
    try {
      setLoader(true);
      setFilter(EMPTY_FILTER); // new level → always wipe filters
      setCurrentFolder(folderPath);
      const base = await getBaseListForLevel(folderPath);
      setDisplayItems(applyFilter(base, EMPTY_FILTER));
    } catch (error) {
      console.error("navigateToFolder error:", error);
    } finally {
      setLoader(false);
    }
  };
  // ─── Click a row: go deeper if folder, open in new tab if file ────────────
  const handleFolderClick = (row: ILoanRecord): void => {
    if (row.folderType === 1) {
      navigateToFolder(row.fileRef);
    } else {
      window.open(row.serverRelativeUrl, "_blank");
    }
  };
  // ─── Breadcrumb: "CSP Loan Files" (Dashboard / root) ───────────────────────
  const goToRoot = (): void => {
    navigateToFolder("");
  };
  // ─── Breadcrumb: any intermediate segment, at any depth ───────────────────
  const goToFolder = (folderPath: string): void => {
    navigateToFolder(folderPath);
  };
  // ─── Search input — filters ONLY the current level's cached base list ─────
  const onSearch = async (value: string): Promise<void> => {
    const newFilter = { ...filter, search: value };
    setFilter(newFilter);

    const base = await getBaseListForLevel(currentFolder);
    setDisplayItems(applyFilter(base, newFilter));
  };
  // ─── Sponsor dropdown — filters ONLY the current level's cached base list ──
  const onSponsorChange = async (
    value: IDrpdownOptions | null,
  ): Promise<void> => {
    const newFilter = { ...filter, sponsor: value };
    setFilter(newFilter);

    const base = await getBaseListForLevel(currentFolder);
    setDisplayItems(applyFilter(base, newFilter));
  };
  // ─── Reset — restore current level's full list from cache, no refetch ─────
  const onReset = async (): Promise<void> => {
    setFilter(EMPTY_FILTER);

    const base = await getBaseListForLevel(currentFolder);
    setDisplayItems(applyFilter(base, EMPTY_FILTER));
  };
  // ─── Build breadcrumb segments from current folder path (any depth) ───────
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
  // ─── Get first + last initials from a full name ────────────────────────────
  const getInitials = (name: string): string => {
    if (!name) return "?";
    const parts = name.trim().split(" ");
    return parts.length >= 2
      ? `${parts[0][0]}${parts[parts.length - 1][0]}`.toUpperCase()
      : name.slice(0, 2).toUpperCase();
  };
  // ─── Column: folder/file name with icon — click to navigate or open ───────
  const nameBodyTemplate = (row: ILoanRecord): React.ReactElement => (
    <span
      className={styles.nameCell}
      onClick={() => handleFolderClick(row)}
      title={row.fileName}
    >
      {row.folderType === 1 ? (
        <i className={`pi pi-folder ${styles.folderIcon}`} />
      ) : (
        <i className={`${getFileIcon(row.fileName)} ${styles.fileIcon}`} />
      )}{" "}
      {row.fileName}
    </span>
  );
  // ─── Column: avatar circle with initials + author name ──────────────────────
  const createdByBodyTemplate = (rowData: ILoanRecord): React.ReactElement => (
    <div className={styles.createdByCell}>
      <span className={styles.avatar} style={{ backgroundColor: AVATAR_COLOR }}>
        {getInitials(rowData.createdby ?? "")}
      </span>
      <span className={styles.createdByName}>{rowData.createdby}</span>
    </div>
  );
  // ─── Format a date value as DD/MM/YYYY ──────────────────────────────────────
  const formatDate = (raw: any): string => {
    if (!raw) return "";
    const d = new Date(raw);
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    return `${dd}/${mm}/${d.getFullYear()}`;
  };
  //date created
  const createdOnBodyTemplate = (rowData: ILoanRecord): React.ReactElement => (
    <span>{formatDate(rowData.createddate)}</span>
  );
  //Date modified
  const modifiedOnBodyTemplate = (rowData: ILoanRecord): React.ReactElement => (
    <span>{formatDate(rowData.modifieddate)}</span>
  );
  //underwritting
  const linkBodyTemplate = (rowData: ILoanRecord) => (
    <span
      className={styles.linkButton}
      onClick={() =>
        navigateToFolder(`${rowData.fileRef}/Asset Management/Underwriting`)
      }
    >
      Underwriting
    </span>
  );
  // ─── Download a folder as a zip (recursive) ─────────────────────────────────
  const downloadFile = async (rowData: any): Promise<void> => {
    try {
      setLoader(true);
      const isFile = rowData.folderType !== 1;
      if (isFile) {
        // ── Direct file download ──────────────────────────────────────
        const buffer = await sp.web
          .getFileByServerRelativePath(rowData.fileRef)
          .getBuffer();

        const blob = new Blob([buffer]);
        const url = window.URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = rowData.fileName;
        anchor.click();
        window.URL.revokeObjectURL(url);
      } else {
        // ── Folder download as zip ────────────────────────────────────
        const folderPath = rowData.fileRef;

        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const JSZipModule = require("jszip");
        const zip = new JSZipModule();

        const addFilesToZip = async (path: string): Promise<void> => {
          const files = await sp.web
            .getFolderByServerRelativePath(path)
            .files();
          for (const file of files) {
            const buffer = await sp.web
              .getFileByServerRelativePath(file.ServerRelativeUrl)
              .getBuffer();
            zip.file(
              file.ServerRelativeUrl.replace(folderPath + "/", ""),
              buffer,
            );
          }
          const subFolders = await sp.web
            .getFolderByServerRelativePath(path)
            .folders();
          for (const subFolder of subFolders) {
            if (subFolder.Name.startsWith("_") || subFolder.Name === "Forms")
              continue;
            zip.file(
              `${subFolder.ServerRelativeUrl.replace(folderPath + "/", "")}/.keep`,
              "",
            );
            await addFilesToZip(subFolder.ServerRelativeUrl);
          }
        };

        await addFilesToZip(folderPath);
        const zipBlob: Blob = await zip.generateAsync({ type: "blob" });
        const url = window.URL.createObjectURL(zipBlob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = `${rowData.fileName}.zip`;
        anchor.click();
        window.URL.revokeObjectURL(url);
      }
    } catch (err) {
      console.error("downloadFile error:", err);
      toastFunc("error", "Error", "Failed to download");
    } finally {
      setLoader(false);
    }
  };
  // ─── Copy SharePoint share link to clipboard ────────────────────────────────
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
          setLoader(false);
          setShowLinkDialog(true);
        } catch {
          const textArea = document.createElement("textarea");
          textArea.value = sharingLink;
          document.body.appendChild(textArea);
          textArea.select();
          document.execCommand("copy");
          document.body.removeChild(textArea);
          setLoader(false);
        }
        return;
      }
    }
  };
  //Rename folder
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
        // ── Step 1: Get list item ID ───────────────────────────────────
        const fileItem: any = await sp.web
          .getFileByServerRelativePath(folderPath)
          .listItemAllFields();

        // ── Step 2: validateUpdateListItem — auto creates version ──────
        await sp.web.lists
          .getByTitle(listNames.loan)
          .items.getById(fileItem.Id)
          .validateUpdateListItem(
            [{ FieldName: "FileLeafRef", FieldValue: fullNewName }],
            false,
          );
      } else {
        // ── Folder: moveByPath + update title ─────────────────────────
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

      // ── Optimistic UI update ───────────────────────────────────────
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
      setDisplayItems(applyFilter(updatedItems, filter));

      if (!currentFolder) {
        setMainFolders(updatedItems);
        dispatch(setLoanDetails(updatedItems));
      }

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
      toastFunc("error", "Error", "Failed to rename");
    } finally {
      setLoader(false);
    }
  };
  //open in app
  const openFileInDesktopApp = (row: any) => {
    const fileUrl = `${window.location.origin}${row.serverRelativeUrl}`;
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
  // ✅ Get version history
  const getVersionHistory = async (rowData: ILoanRecord): Promise<void> => {
    try {
      setLoader(true);

      const file = sp.web.getFileByServerRelativePath(rowData.fileRef);
      const item = await file.getItem();

      // ── Step 1: Get ALL fields without expand to find internal names
      const allFields = await item.select("*")();

      const [fileInfo, versions] = await Promise.all([
        item
          .select(
            "FileLeafRef",
            "Modified",
            "OData__UIVersionString",
            "Author/Title",
            "Editor/Title",
            "SponsorId",
          )
          .expand("Author", "Editor")(),
        file.versions(),
      ]);

      // ── Fetch sponsor name ─────────────────────────────────────────
      let sponsorName = "";
      if (fileInfo.SponsorId) {
        try {
          const sponsorItem = await sp.web.lists
            .getByTitle(listNames.sponsors)
            .items.getById(fileInfo.SponsorId)
            .select("Title")();
          sponsorName = sponsorItem?.Title || "";
        } catch {
          sponsorName = "";
        }
      }

      const getMetadata = (field: any): string => {
        if (!field) return "";
        if (field?.Label) return field.Label;
        if (Array.isArray(field) && field.length) return field[0]?.Label || "";
        if (typeof field === "string") return field;
        return "";
      };

      const history: any[] = [];

      history.push({
        versionLabel: fileInfo.OData__UIVersionString || "",
        modified: fileInfo.Modified
          ? new Date(fileInfo.Modified).toLocaleString()
          : "",
        modifiedBy: fileInfo.Editor?.Title || "",
        size: "-",
        comments: "Current Version",
        name: fileInfo.FileLeafRef || "",
        sponsor: sponsorName,
        createdby: fileInfo.Author?.Title || "",
        createddate: fileInfo.Modified
          ? new Date(fileInfo.Modified).toLocaleString()
          : "",
        // ── allFields has correct internal names ───────────────────
        assetmanagement: getMetadata(allFields.AssetManagement),
        servicing: getMetadata(allFields.Servicing),
        legal: getMetadata(allFields.Legal),
      });

      for (const version of versions) {
        history.push({
          versionLabel: version.VersionLabel || "",
          modified: version.Created
            ? new Date(version.Created).toLocaleString()
            : "",
          modifiedBy: version.CreatedBy?.Title || "",
          size: version.Size ? `${(version.Size / 1024).toFixed(2)} KB` : "",
          comments: version.CheckInComment || "",
          name: fileInfo.FileLeafRef || "",
          sponsor: sponsorName,
          createdby: fileInfo.Author?.Title || "",
          createddate: version.Created
            ? new Date(version.Created).toLocaleString()
            : "",
          assetmanagement: getMetadata(allFields.AssetManagement),
          servicing: getMetadata(allFields.Servicing),
          legal: getMetadata(allFields.Legal),
        });
      }

      history.sort(
        (a, b) => parseFloat(b.versionLabel) - parseFloat(a.versionLabel),
      );

      setVersionHistory(history);
      setShowVersionHistory(true);
    } catch (error) {
      console.error("Version History Error:", error);
      toastFunc("error", "Error", "Failed to load version history");
    } finally {
      setLoader(false);
    }
  };
  // ─── Action column — opens the popup menu for THIS row only ────────────────
  const actionTemplate = (rowData: any): React.ReactElement => (
    <Button
      icon="pi pi-ellipsis-h"
      text
      rounded
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
      setLoader(true);
      const trimmedName = folderName.trim();
      const folderPath = `${currentFolder}/${trimmedName}`;

      // ── Get sponsor from root loan folder (already in state) ───────
      const rootLoanFolder = mainFolders.find((f) =>
        currentFolder.includes(f.fileRef),
      );
      const sponsorId = rootLoanFolder?.sponsor?.id ?? null;

      await sp.web.folders.addUsingPath(folderPath);

      // ── Set sponsor lookup in one call ─────────────────────────────
      if (sponsorId) {
        const folderItem: any = await sp.web
          .getFolderByServerRelativePath(folderPath)
          .listItemAllFields();

        await sp.web.lists
          .getByTitle(listNames.loan)
          .items.getById(folderItem.Id)
          .update({ SponsorId: sponsorId });
      }

      toastFunc("success", "Success", "Folder created successfully");
      setFolderName("");
      setShowCreateFolder(false);

      const updatedItems = await fetchFolderContents(currentFolder);
      folderCacheRef.current.set(currentFolder, updatedItems);
      setDisplayItems(applyFilter(updatedItems, filter));
    } catch (error) {
      console.error("Create Folder Error:", error);
      toastFunc("error", "Error", "Failed to create folder");
    } finally {
      setLoader(false);
    }
  };
  // share functionality
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
        `&ma=0`;

      const width = 550;
      const height = 400;

      const left = window.screenX + (window.outerWidth - width) / 2;
      const top = window.screenY + (window.outerHeight - height) / 2;

      const features = [
        `width=${width}`,
        `height=${height}`,
        `left=${left}`,
        `top=${top}`,
        "resizable=yes",
        "scrollbars=yes",
        "toolbar=no",
        "menubar=no",
        "location=no",
        "status=no",
      ].join(",");

      window.open(shareUrl, "ShareDialog", features);
    } catch (error) {
      console.error("Open Share Dialog Error:", error);
    }
  };
  // copy move functionality
  const openCopyMoveDialog = (
    rowData: ILoanRecord,
    operation: "copy" | "move",
  ): void => {
    setCopyMoveOperation(operation);
    setCopyMoveTarget(rowData);
    setSelectedDestination(null);
    setShowCopyMoveDialog(true);
  };
  const executeCopyMove = async (): Promise<void> => {
    if (!selectedDestination) {
      toastFunc("warn", "Warning", "Please select a destination folder");
      return;
    }

    try {
      setLoader(true);
      const sourcePath = copyMoveTarget?.fileRef || "";
      const fileName = copyMoveTarget?.fileName || "";
      const destinationUrl = selectedDestination.ServerRelativeUrl;

      if (!destinationUrl) {
        toastFunc("warn", "Warning", "Selected destination folder is invalid.");
        return;
      }

      const isFile = copyMoveTarget?.folderType !== 1;
      const destinationPath = `${destinationUrl}/${fileName}`;

      if (copyMoveOperation === "copy") {
        if (isFile) {
          await sp.web
            .getFileByServerRelativePath(sourcePath)
            .copyByPath(destinationPath, true, true);
        } else {
          await sp.web
            .getFolderByServerRelativePath(sourcePath)
            .copyByPath(destinationPath, true);
        }
      } else {
        if (isFile) {
          await sp.web
            .getFileByServerRelativePath(sourcePath)
            .moveByPath(destinationPath, true);
        } else {
          await sp.web
            .getFolderByServerRelativePath(sourcePath)
            .moveByPath(destinationPath, true);
        }
      }

      // ── Cache invalidation ─────────────────────────────────────────
      folderCacheRef.current.delete(sourcePath);
      folderCacheRef.current.delete(destinationUrl);

      if (copyMoveOperation === "move") {
        const levelPath =
          currentFolder || sourcePath.substring(0, sourcePath.lastIndexOf("/"));
        const cachedItems =
          folderCacheRef.current.get(levelPath) ?? mainFolders;
        const updatedItems = cachedItems.filter(
          (item) => item.fileRef !== sourcePath,
        );
        folderCacheRef.current.set(levelPath, updatedItems);
        setDisplayItems(applyFilter(updatedItems, filter));

        if (!currentFolder) {
          setMainFolders(updatedItems);
          dispatch(setLoanDetails(updatedItems));
        }
      }

      toastFunc(
        "success",
        "Success",
        `${isFile ? "File" : "Folder"} ${copyMoveOperation === "copy" ? "copied" : "moved"} successfully`,
      );
      setShowCopyMoveDialog(false);
      setCopyMoveTarget(null);
      setSelectedDestination(null);
    } catch (error) {
      console.error("executeCopyMove error:", error);
      toastFunc(
        "error",
        "Error",
        `Failed to ${copyMoveOperation === "copy" ? "copy" : "move"}`,
      );
    } finally {
      setLoader(false);
    }
  };

  // ─── Load data on mount ──────────────────────────────────────────────────────
  useEffect(() => {
    const handleOutsideClick = (event: MouseEvent) => {
      menuRef.current?.hide(event as any);
    };
    document.addEventListener("click", handleOutsideClick);
    setLoader(true);
    getSponsorData().catch((error) => console.error(error));
    return () => document.removeEventListener("click", handleOutsideClick);
  }, []);

  return (
    <>
      {loader ? (
        <Loader />
      ) : (
        <div className={styles.container}>
          {/* ── Toolbar: breadcrumb, search, sponsor filter, refresh, new ── */}
          <div className={styles.toolbar}>
            <div className={styles.toolbarTitle}>
              <span className={styles.breadcrumbRoot} onClick={goToRoot}>
                CSP Loan Files
              </span>
              {getBreadcrumbSegments().map((segment, index, arr) => (
                <React.Fragment key={segment.path}>
                  <span className={styles.breadcrumbSeparator}>&gt;&gt;</span>
                  <span
                    className={
                      index === arr.length - 1
                        ? styles.breadcrumbCurrent
                        : styles.breadcrumbLink
                    }
                    onClick={() =>
                      index < arr.length - 1 && goToFolder(segment.path)
                    }
                  >
                    {segment.label}
                  </span>
                </React.Fragment>
              ))}
            </div>

            <div className={styles.toolbarSearch}>
              <i className={`pi pi-search ${styles.searchIcon}`} />
              <InputText
                placeholder="Search loan..."
                className={styles.searchInput}
                value={filter.search}
                onChange={(e) => onSearch(e.target.value)}
              />
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
            {currentFolder && (
              <>
                <Button
                  className={styles.newBtn}
                  onClick={(e) => {
                    getNewMenuItem();
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
                  model={menuItems}
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
              tableStyle={{ minWidth: "100%" }}
              paginatorTemplate="CurrentPageReport RowsPerPageDropdown FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink"
              currentPageReportTemplate="Showing {first} to {last} of {totalRecords} records"
              className={styles.table}
            >
              <Column
                field="fileName"
                header="Name"
                body={nameBodyTemplate}
                sortable
                style={{ width: "22%" }}
              />
              <Column
                field="sponsor.sponsorTitle"
                header="Sponsor"
                sortable
                style={{ width: "15%" }}
              />
              {currentFolder.includes("Asset Management") && (
                <Column header="Asset Mgmt" style={{ width: "15%" }} />
              )}
              {currentFolder.includes("CSP Legal") && (
                <Column header="Legal" style={{ width: "15%" }} />
              )}
              {currentFolder.includes("Servicing") && (
                <Column header="Servicing" style={{ width: "15%" }} />
              )}
              <Column
                header="Created By"
                body={createdByBodyTemplate}
                sortable
                style={{ width: "18%" }}
              />
              <Column
                header="Created On"
                body={createdOnBodyTemplate}
                sortable
                style={{ width: "10%" }}
              />
              <Column
                header="Modified On"
                body={modifiedOnBodyTemplate}
                sortable
                style={{ width: "10%" }}
              />
              {!currentFolder && (
                <Column
                  header="Link"
                  body={linkBodyTemplate}
                  style={{ width: "10%" }}
                />
              )}

              <Column
                header="Action"
                style={{ width: "8%" }}
                body={actionTemplate}
              />
            </DataTable>
          </div>
          {/* Rename Dialog */}
          <Dialog
            visible={showRename}
            style={{ width: "35%" }}
            className="dialogCloseIcon"
            onHide={() => setShowRename(false)}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
          >
            <h3 className="modelHeader" style={{ marginBottom: "10px" }}>
              Rename
            </h3>
            <div className={styles.modalBodyFlex}>
              <div className={styles.fields}>
                <div
                  style={{ display: "flex", alignItems: "center", gap: "8px" }}
                >
                  <InputText
                    className="singlelineText"
                    placeholder="Enter folder name"
                    value={folderName}
                    style={{ flex: 1 }}
                    onChange={(e) => setFolderName(e.target.value)}
                  />
                  {/* Show extension only for files, not folders */}
                  {selectedFolder?.folderType !== 1 && (
                    <span
                      style={{
                        whiteSpace: "nowrap",
                        color: "#6c757d",
                        fontWeight: 500,
                      }}
                    >
                      {`.${selectedFolder?.fileName?.split(".").pop()}`}
                    </span>
                  )}
                </div>
              </div>
            </div>
            <div className={styles.modalFooter}>
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
          {/* copy link dialog */}
          <Dialog
            visible={showLinkDialog}
            onHide={() => setShowLinkDialog(false)}
            showHeader={false}
            draggable={false}
            resizable={false}
            style={{ width: "32rem" }}
            className={styles.linkDialog}
          >
            <div className={styles.dialogHeader}>
              <div className={styles.successIcon}>
                <i className="pi pi-check" />
              </div>

              <div className={styles.dialogContent}>
                <h2>Link copied</h2>
              </div>

              <i
                className={`pi pi-times ${styles.closeIcon}`}
                onClick={() => setShowLinkDialog(false)}
              />
            </div>
          </Dialog>
          {/* ── Version History Dialog ───────────────────────────── */}
          <Dialog
            visible={showVersionHistory}
            style={{ width: "90%" }}
            header={
              <div>
                <h3 style={{ margin: 0 }}>Version History</h3>
                <p
                  style={{
                    margin: "4px 0 0",
                    fontSize: "13px",
                    color: "#6b7280",
                  }}
                >
                  View all available versions of this document
                </p>
              </div>
            }
            onHide={() => setShowVersionHistory(false)}
            draggable={false}
          >
            <DataTable
              value={versionHistory}
              scrollable
              scrollHeight="400px"
              className={styles.table}
              emptyMessage="No version history found"
              stripedRows
            >
              <Column
                field="versionLabel"
                header="Version"
                style={{ width: "6%" }}
                sortable
              />
              <Column
                field="name"
                header="Name"
                style={{ width: "12%" }}
                sortable
              />
              <Column
                field="sponsor"
                header="Sponsor"
                style={{ width: "10%" }}
                sortable
              />
              <Column
                field="createdBy"
                header="Created By"
                style={{ width: "10%" }}
                sortable
              />
              <Column
                field="createddate"
                header="Created Date"
                style={{ width: "10%" }}
                sortable
              />
              <Column
                field="modified"
                header="Modified Date"
                style={{ width: "10%" }}
                sortable
              />
              <Column
                field="modifiedBy"
                header="Modified By"
                style={{ width: "10%" }}
                sortable
              />
              <Column
                field="assetmanagement"
                header="Asset Management"
                style={{ width: "10%" }}
                sortable
              />
              <Column
                field="servicing"
                header="Servicing"
                style={{ width: "8%" }}
                sortable
              />
              <Column
                field="legal"
                header="Legal"
                style={{ width: "8%" }}
                sortable
              />
              <Column
                field="size"
                header="Size"
                style={{ width: "6%" }}
                sortable
              />
              <Column
                field="comments"
                header="Comments"
                style={{ width: "10%" }}
                body={(row) =>
                  row.comments === "Current Version" ? (
                    <span
                      style={{
                        background: "#ede9fe",
                        color: "#4f46e5",
                        padding: "2px 8px",
                        borderRadius: "12px",
                        fontSize: "12px",
                        fontWeight: 600,
                      }}
                    >
                      Current Version
                    </span>
                  ) : (
                    <span>{row.comments || "-"}</span>
                  )
                }
              />
            </DataTable>

            <div className={styles.modalFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Close"
                onClick={() => setShowVersionHistory(false)}
              />
            </div>
          </Dialog>
          {/* subfolder dialog */}
          <Dialog
            visible={showCreateFolder}
            style={{ width: "35%" }}
            className="dialogCloseIcon"
            onHide={() => setShowCreateFolder(false)}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
          >
            <h3 className="modelHeader">Create Sub-Folder</h3>
            <div className={styles.modalBodyFlex}>
              <div className={styles.fields}>
                <label>
                  Folder Name <span className={styles.required}>*</span>
                </label>
                <InputText
                  className="singlelineText"
                  placeholder="Enter sub folder name"
                  value={folderName}
                  onChange={(e) => setFolderName(e.target.value)}
                />
              </div>
            </div>
            <div className={styles.modalFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                onClick={() => setShowCreateFolder(false)}
              />
              <Button
                label="Create"
                icon="pi pi-plus"
                className="submitBtn"
                onClick={createsubFolder}
              />
            </div>
          </Dialog>

          {/* ── Delete confirmation dialog ── */}
          <Dialog
            visible={showDeleteDialog}
            style={{ width: "30%" }}
            className="dialogCloseIcon"
            onHide={() => {
              console.log("");
            }}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
          >
            <h3 className="modelHeader">Delete Loan</h3>
            <div className={styles.dialogbody}>
              <div className={styles.innerDiv}>
                <i className="pi pi-trash" />
                <span>Are you sure you want to delete {}?</span>
              </div>
            </div>
            <div className={styles.modalFooter}>
              <Button
                className="cancelBtn"
                label="No"
                icon="pi pi-times"
                onClick={() => setshowDeleteDialog(false)}
              />
              <Button
                label="Yes"
                icon="pi pi-trash"
                className="submitBtn"
                onClick={() => {
                  setLoader(true);
                }}
              />
            </div>
          </Dialog>
          {/* ── Copy To / Move To Dialog ── */}
          <Dialog
            visible={showCopyMoveDialog}
            style={{ width: "45%" }}
            header={`${copyMoveOperation === "copy" ? "Copy" : "Move"} "${copyMoveTarget?.fileName}" to...`}
            onHide={() => {
              setShowCopyMoveDialog(false);
              setCopyMoveTarget(null);
              setSelectedDestination(null);
            }}
            draggable={false}
          >
            <FolderPicker
              context={props.context as any}
              label="Select Destination Folder"
              rootFolder={{
                Name: listNames.loan,
                ServerRelativeUrl: `${props.context.pageContext.web.serverRelativeUrl}/${listNames.loan}`,
              }}
              defaultFolder={selectedDestination as any}
              onSelect={(folder: IFolder) => setSelectedDestination(folder)}
              canCreateFolders={false}
            />

            {selectedDestination && (
              <p
                style={{ fontSize: "13px", color: "#4f46e5", marginTop: "8px" }}
              >
                <i className="pi pi-folder" style={{ marginRight: "6px" }} />
                Destination: <strong>{selectedDestination.Name}</strong>
              </p>
            )}

            <div className={styles.modalFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                onClick={() => {
                  setShowCopyMoveDialog(false);
                  setCopyMoveTarget(null);
                  setSelectedDestination(null);
                }}
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
          <Menu model={menuItems} popup ref={menuRef} />
        </div>
      )}
    </>
  );
};

export default Loan;
