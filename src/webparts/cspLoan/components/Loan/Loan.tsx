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
import {
  folderStructure,
  listNames,
  toastFunc,
} from "../../assets/Config/Config";
import styles from "./Loan.module.scss";

const AVATAR_COLOR = "#4f46e5";
const LIBRARY_NAME = "exchange";

interface IAddLoan {
  loannumber: string;
  sponsor: { code: number | null; name: string };
}

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
    id: Number(item.ID),
    sponsorTitle: item.Sponsor || "",
  },
  createdby: item.Author?.[0]?.title || item.Author?.title || "",
  createddate: item.Created || null,
  modifieddate: item.Modified || null,
  assetmanagement: "",
  servicing: "",
  legal: "",
});

const Loan = (props: IProps) => {
  const initialLoanState: IAddLoan = {
    loannumber: "",
    sponsor: { code: null, name: "" },
  };

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

  const [showDialog, setShowDialog] = useState(false);
  const [addLoan, setAddLoan] = useState<IAddLoan>({ ...initialLoanState });
  const [menuItems, setMenuItems] = useState<any[]>([]);

  // ─── "New" dropdown menu items ─────────────────────────────────────────────
  const getNewMenuItem = (): void => {
    setMenuItems([
      {
        label: "New Loan",
        icon: "pi pi-plus",
        visible: !currentFolder,
        command: () => setShowDialog(true),
      },
      {
        label: "New sub folder",
        icon: "pi pi-folder-plus",
        visible: !!currentFolder,
        command: () => console.log("New sub folder clicked at:", currentFolder),
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
        command: () => console.log(selectedRowRef.current),
      },
      {
        label: "Download",
        icon: "pi pi-download",
        command: () => downloadFile(selectedRowRef.current),
      },
      { label: "Share", icon: "pi pi-share-alt", command: () => {} },
      {
        label: "Copy Link",
        icon: "pi pi-link",
        command: () => copySharePointLink(selectedRowRef.current),
      },
      {
        label: "Move To",
        icon: "pi pi-arrow-right-arrow-left",
        command: () => {},
      },
      { label: "Copy To", icon: "pi pi-copy", command: () => {} },
      { label: "Rename", icon: "pi pi-pencil", command: () => {} },
      {
        label: "Version History",
        icon: "pi pi-history",
        visible: selectedRowRef.current?.folderType !== 1,
        command: () => {},
      },
      {
        label: "Delete",
        icon: "pi pi-trash",
        className: "deleteMenu",
        command: () => {},
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
    <span className={styles.nameCell} onClick={() => handleFolderClick(row)}>
      {row.folderType === 1 ? "📁" : "📄"} {row.fileName}
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

  const createdOnBodyTemplate = (rowData: ILoanRecord): React.ReactElement => (
    <span>{formatDate(rowData.createddate)}</span>
  );
  const modifiedOnBodyTemplate = (rowData: ILoanRecord): React.ReactElement => (
    <span>{formatDate(rowData.modifieddate)}</span>
  );

  // ─── Update dialog state on field change ───────────────────────────────────
  const onChangeHandler = <K extends keyof IAddLoan>(
    key: K,
    value: IAddLoan[K],
  ): void => {
    let addValue: any = value;
    if (key === "sponsor" && value) {
      addValue = { code: (value as any).code, name: (value as any).name };
    }
    setAddLoan((prev) => ({ ...prev, [key]: addValue }));
  };

  // ─── Create subfolder structure recursively ────────────────────────────────
  const createFolderStructure = async (
    parentPath: string,
    folders: any[],
    sponsorId: number | null,
  ): Promise<void> => {
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    for (const folder of folders) {
      const folderPath = `${parentPath}/${folder.name}`;

      await sp.web.folders.addUsingPath(folderPath);

      const folderItem = await sp.web
        .getFolderByServerRelativePath(`${siteUrl}/${folderPath}`)
        .listItemAllFields();

      await sp.web.lists
        .getByTitle(listNames.loan)
        .items.getById(folderItem.Id)
        .update({ SponsorId: sponsorId });

      if (folder.children?.length) {
        await createFolderStructure(folderPath, folder.children, sponsorId);
      }
    }
  };

  // ─── Create a brand-new Loan Number folder + its standard subfolder tree ──
  const createLoanStructure = async (
    loanNumber: string,
    sponsorId: number | null,
  ): Promise<void> => {
    try {
      setLoader(true);
      const libraryName = listNames.loan;
      const rootPath = `${libraryName}/${loanNumber}`;
      const siteUrl = props.context.pageContext.web.absoluteUrl;

      await sp.web.folders.addUsingPath(rootPath);
      const mainFolderItem = await sp.web
        .getFolderByServerRelativePath(`${siteUrl}/${rootPath}`)
        .listItemAllFields();

      await sp.web.lists
        .getByTitle(libraryName)
        .items.getById(mainFolderItem.Id)
        .update({ SponsorId: sponsorId });

      await createFolderStructure(rootPath, folderStructure, sponsorId);

      toastFunc(
        "success",
        "Success",
        `Loan Number ${loanNumber} created successfully`,
      );

      // New root folder created — refresh root list + clear cache
      await getSponsorData();
    } catch (error) {
      console.error(error);
    } finally {
      setLoader(false);
    }
  };

  // ─── Validate New Loan form ──────────────────────────────────────────────────
  const loanValidation = (): void => {
    if (!addLoan.loannumber.trim()) {
      toastFunc("warn", "Warning", "Please enter loan number");
    } else if (!addLoan.sponsor.name) {
      toastFunc("warn", "Warning", "Please select sponsor");
    } else {
      setShowDialog(false);
      createLoanStructure(addLoan.loannumber, addLoan.sponsor.code);
    }
  };

  // ─── Download a folder as a zip (recursive) ─────────────────────────────────
  const downloadFile = async (rowData: any): Promise<void> => {
    setLoader(true);
    const folderPath = rowData.fileRef;
    const folderName = rowData.fileName;

    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const JSZipModule = require("jszip");
    const zip = new JSZipModule();

    const addFilesToZip = async (path: string): Promise<void> => {
      const files = await sp.web.getFolderByServerRelativePath(path).files();
      for (const file of files) {
        const buffer = await sp.web
          .getFileByServerRelativePath(file.ServerRelativeUrl)
          .getBuffer();
        const relativePath = file.ServerRelativeUrl.replace(
          folderPath + "/",
          "",
        );
        zip.file(relativePath, buffer);
      }

      const subFolders = await sp.web
        .getFolderByServerRelativePath(path)
        .folders();
      for (const subFolder of subFolders) {
        if (subFolder.Name.startsWith("_") || subFolder.Name === "Forms")
          continue;

        const relativeFolderPath = subFolder.ServerRelativeUrl.replace(
          folderPath + "/",
          "",
        );
        zip.file(`${relativeFolderPath}/.keep`, "");
        await addFilesToZip(subFolder.ServerRelativeUrl);
      }
    };

    try {
      await addFilesToZip(folderPath);
      const zipBlob: Blob = await zip.generateAsync({ type: "blob" });
      const url = window.URL.createObjectURL(zipBlob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `${folderName}.zip`;
      anchor.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("downloadFile error:", err);
    } finally {
      setLoader(false);
    }
  };

  // ─── Copy SharePoint share link to clipboard ────────────────────────────────
  const copySharePointLink = async (rowData: any): Promise<void> => {
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
          toastFunc("success", "Success", "Folder link copied!");
        } catch {
          const textArea = document.createElement("textarea");
          textArea.value = sharingLink;
          document.body.appendChild(textArea);
          textArea.select();
          document.execCommand("copy");
          document.body.removeChild(textArea);
        }
        return;
      }
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
                  <span className={styles.breadcrumbSeparator}>&gt;</span>
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
                style={{ width: "16%" }}
              />
              <Column
                field="sponsor.sponsorTitle"
                header="Sponsor"
                sortable
                style={{ width: "16%" }}
              />
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
                style={{ width: "12%" }}
              />
              <Column
                header="Modified On"
                body={modifiedOnBodyTemplate}
                sortable
                style={{ width: "12%" }}
              />
              <Column
                header="Action"
                style={{ width: "9%" }}
                body={actionTemplate}
              />
            </DataTable>
          </div>

          {/* ── Add New Loan dialog ── */}
          <Dialog
            visible={showDialog}
            style={{ width: "35%" }}
            className="dialogCloseIcon"
            onHide={() => setAddLoan(initialLoanState)}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
          >
            <h3 className="modelHeader">Add Loan Number</h3>
            <div className={styles.modalBodyFlex}>
              <div className={styles.fields}>
                <label>
                  Loan Number <span className={styles.required}>*</span>
                </label>
                <InputText
                  className="singlelineText"
                  placeholder="Enter loan number"
                  value={addLoan.loannumber}
                  onChange={(e) =>
                    onChangeHandler("loannumber", e.target.value)
                  }
                />
              </div>
              <div className={styles.fields}>
                <label>
                  Sponsor <span className={styles.required}>*</span>
                </label>
                <Dropdown
                  options={drpdown.sponsor}
                  optionLabel="name"
                  placeholder="Select sponsor"
                  value={addLoan.sponsor}
                  filter
                  onChange={(e) => onChangeHandler("sponsor", e.value)}
                />
              </div>
            </div>
            <div className={styles.modalFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                onClick={() => {
                  setAddLoan(initialLoanState);
                  setShowDialog(false);
                }}
              />
              <Button
                label="Create"
                icon="pi pi-plus"
                className="submitBtn"
                onClick={loanValidation}
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
