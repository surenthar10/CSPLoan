/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-use-before-define */
import { sp } from "@pnp/sp/presets/all";
import * as React from "react";
import { useEffect, useState } from "react";
import { IDialogDetails, ISponsorRecord } from "../../assets/Config/interface";
import {
  actions,
  flags,
  listNames,
  toastFunc,
} from "../../assets/Config/Config";
import { InputText } from "primereact/inputtext";
import { Button } from "primereact/button";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
import { Dialog } from "primereact/dialog";
import { InputTextarea } from "primereact/inputtextarea";
import { addItem, updateItem } from "../../assets/Config/function";
import Loader from "../Loader/Loader";
import styles from "./Sponsor.module.scss";

// ─── Avatar color palette (one color for all avatars) ─────────────────────────
const AVATAR_COLOR = "#4f46e5";

// ─── Default empty state values ───────────────────────────────────────────────
const EMPTY_SPONSOR: ISponsorRecord = {
  id: null,
  sponsor: "",
  description: "",
  loans: [],
  createdby: "",
  date: null,
};

const EMPTY_DIALOG: IDialogDetails = {
  condition: false,
  data: EMPTY_SPONSOR,
  type: flags.empty,
};

const Sponsor = () => {
  const [sponsorList, setSponsorList] = useState<ISponsorRecord[]>([]);
  const [sponsorFilterData, setSponsorFilterData] = useState<ISponsorRecord[]>(
    [],
  );
  const [searchText, setSearchText] = useState("");
  const [loader, setLoader] = useState(false);
  const [dialog, setDialog] = useState<IDialogDetails>(EMPTY_DIALOG);

  // ─── Fetch all sponsors from SharePoint ───────────────────────────────────
  const getSponsorData = async (type: string): Promise<void> => {
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

      const sponsorArray: ISponsorRecord[] = allRows.map((item: any) => ({
        id: item.Id,
        sponsor: item.Title,
        description: item.Description,
        loans: [],
        createdby: item.Author?.Title,
        date: item.Created,
      }));

      setSponsorList(sponsorArray);
      setSponsorFilterData(sponsorArray);
      setLoader(false);

      if (type === "New")
        toastFunc("success", "Success", "Sponsor created successfully");
      if (type === "Update")
        toastFunc("success", "Success", "Sponsor updated successfully");
      if (type === "Delete")
        toastFunc("success", "Success", "Sponsor deleted successfully");
    } catch (error) {
      console.error("getSponsorData error:", error);
    }
  };

  // ─── Search sponsors by name or description ────────────────────────────────
  const onSearch = (value: string): void => {
    setSearchText(value);
    const filtered = sponsorList.filter(
      (item) =>
        item.sponsor?.toLowerCase().includes(value.toLowerCase()) ||
        item.description?.toLowerCase().includes(value.toLowerCase()),
    );
    setSponsorFilterData(filtered);
  };

  // ─── Reset search and reload sponsor list ─────────────────────────────────
  const refreshData = async (type: string): Promise<void> => {
    setSearchText("");
    await getSponsorData(type);
  };

  // ─── Update dialog state on field change ──────────────────────────────────
  const onChangeHandler = <K extends keyof ISponsorRecord>(
    key: K,
    value: ISponsorRecord[K],
  ): void => {
    setDialog((prev) => ({
      ...prev,
      data: { ...prev.data, [key]: value },
    }));
  };

  // ─── Add new sponsor to SharePoint list ───────────────────────────────────
  const addSponsor = async (addData: IDialogDetails): Promise<void> => {
    const payload = {
      Title: addData.data.sponsor,
      Description: addData.data.description,
    };
    const response = await addItem(listNames.sponsors, payload, "addSponsor");
    if (response) {
      setDialog(EMPTY_DIALOG);
      await refreshData("New");
    }
  };

  // ─── Update or soft-delete an existing sponsor ────────────────────────────
  const updateSponsor = async (
    addData: IDialogDetails,
    isDelete: boolean,
  ): Promise<void> => {
    const payload = {
      Title: addData.data.sponsor,
      Description: addData.data.description,
      IsDelete: isDelete,
    };
    const response = await updateItem(
      listNames.sponsors,
      payload,
      addData.data.id,
      "UpdateSponsor",
    );
    if (response) {
      setDialog(EMPTY_DIALOG);
      await refreshData(isDelete ? "Delete" : "Update");
    }
  };

  // ─── Validate form and trigger add or update ──────────────────────────────
  const sponsorValidation = async (): Promise<void> => {
    if (!dialog.data.sponsor.trim()) {
      toastFunc("warn", "Warning", "Please enter sponsor");
      return;
    }
    setLoader(true);
    if (dialog.type === flags.add) {
      await addSponsor(dialog);
    } else {
      await updateSponsor(dialog, false);
    }
  };

  // ─── Get first + last initials from a full name ───────────────────────────
  const getInitials = (name: string): string => {
    if (!name) return "?";
    const parts = name.trim().split(" ");
    return parts.length >= 2
      ? `${parts[0][0]}${parts[parts.length - 1][0]}`.toUpperCase()
      : name.slice(0, 2).toUpperCase();
  };

  // ─── Column: truncated description with tooltip ───────────────────────────
  const descriptionTemplate = (rowData: ISponsorRecord): React.ReactElement => (
    <div
      title={rowData.description}
      style={{
        maxWidth: "350px",
        overflow: "hidden",
        whiteSpace: "nowrap",
        textOverflow: "ellipsis",
        cursor: "pointer",
      }}
    >
      {rowData.description}
    </div>
  );

  // ─── Column: avatar circle with initials + author name ────────────────────
  const createdByBodyTemplate = (
    rowData: ISponsorRecord,
  ): React.ReactElement => {
    const name = rowData.createdby ?? "";
    return (
      <div className={styles.createdByCell}>
        <span
          className={styles.avatar}
          style={{ backgroundColor: AVATAR_COLOR }}
        >
          {getInitials(name)}
        </span>
        <span className={styles.createdByName}>{name}</span>
      </div>
    );
  };

  // ─── Column: format date as DD/MM/YYYY ────────────────────────────────────
  const createdOnBodyTemplate = (
    rowData: ISponsorRecord,
  ): React.ReactElement => {
    let display = "";
    if (rowData.date) {
      const d = new Date(rowData.date);
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const yyyy = d.getFullYear();
      display = `${dd}/${mm}/${yyyy}`;
    }
    return <span>{display}</span>;
  };

  // ─── Column: edit and delete action icons ─────────────────────────────────
  const actionTemplate = (rowData: ISponsorRecord): React.ReactElement => (
    <div className={styles.actionIcons}>
      <i
        className="pi pi-pencil"
        onClick={() =>
          setDialog({
            ...dialog,
            condition: true,
            type: flags.edit,
            data: { ...rowData },
          })
        }
      />
      <i
        className="pi pi-trash"
        onClick={() =>
          setDialog({
            ...dialog,
            condition: true,
            type: flags.delete,
            data: { ...rowData },
          })
        }
      />
    </div>
  );

  // ─── Load sponsors on mount ───────────────────────────────────────────────
  useEffect(() => {
    setLoader(true);
    getSponsorData("").catch((error) => console.error(error));
  }, []);

  return (
    <>
      {loader ? (
        <Loader />
      ) : (
        <div className={styles.sponsorContainer}>
          {/* ── Toolbar: title, search, refresh, new sponsor ── */}
          <div className={styles.sponsortoolbar}>
            <div className={styles.toolbarTitle}>Sponsor Details</div>

            <div className={styles.toolbarSearch}>
              <i
                className="pi pi-search"
                style={{
                  position: "absolute",
                  left: "8px",
                  top: "55%",
                  transform: "translateY(-50%)",
                  color: "#9ca3af",
                  fontSize: "14px",
                }}
              />
              <InputText
                placeholder="Search sponsors..."
                className={styles.searchInput}
                value={searchText}
                onChange={(e) => onSearch(e.target.value)}
              />
            </div>

            <Button
              icon="pi pi-refresh"
              className={styles.refreshBtn}
              onClick={() => refreshData("Refresh")}
            />
            <Button
              label="New Sponsor"
              icon="pi pi-plus"
              className={styles.newSponsorBtn}
              onClick={() =>
                setDialog({ ...dialog, condition: true, type: flags.add })
              }
            />
          </div>

          {/* ── Sponsor data table ── */}
          <div className={styles.tableContainer}>
            <DataTable
              value={sponsorFilterData}
              paginator
              rows={10}
              rowsPerPageOptions={[10, 25, 50, 100]}
              emptyMessage="No sponsors found"
              tableStyle={{ minWidth: "100%" }}
              paginatorTemplate="CurrentPageReport RowsPerPageDropdown FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink"
              currentPageReportTemplate="Showing {first} to {last} of {totalRecords} records"
              className={styles.sponsorTable}
            >
              <Column
                field="sponsor"
                header="Sponsor"
                sortable
                style={{ width: "16%" }}
              />
              <Column
                body={descriptionTemplate}
                header="Description"
                sortable
                style={{ width: "23%" }}
              />
              <Column header="Loans" sortable style={{ width: "22%" }} />
              <Column
                header="Created By"
                sortable
                style={{ width: "18%" }}
                body={createdByBodyTemplate}
              />
              <Column
                header="Created On"
                sortable
                style={{ width: "12%" }}
                body={createdOnBodyTemplate}
              />
              <Column
                header="Action"
                style={{ width: "9%" }}
                body={actionTemplate}
              />
            </DataTable>
          </div>

          {/* ── Add / Edit dialog ── */}
          <Dialog
            visible={dialog.condition && dialog.type !== flags.delete}
            style={{ width: "35%" }}
            className="dialogCloseIcon"
            onHide={() => {
              console.log("");
            }}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
          >
            <h3 className="modelHeader">
              {dialog.type === flags.add ? actions.add : actions.edit} Sponsor
            </h3>
            <h5 className="modelsubHeader">
              {dialog.type === flags.add
                ? "Create a new sponsor and associate loan numbers"
                : "Update details and manage loan associations"}
            </h5>
            <div className={styles.modalBodyFlex}>
              <div className={styles.fields}>
                <label>
                  Sponsor <span className={styles.required}>*</span>
                </label>
                <InputText
                  className="singlelineText"
                  placeholder="Enter sponsor"
                  value={dialog.data.sponsor}
                  onChange={(e) => onChangeHandler("sponsor", e.target.value)}
                />
              </div>
              <div className={styles.fields}>
                <label>Description</label>
                <InputTextarea
                  className="singlelineText"
                  placeholder="Enter description"
                  value={dialog.data.description}
                  style={{ height: "80px" }}
                  onChange={(e) =>
                    onChangeHandler("description", e.target.value)
                  }
                />
              </div>
            </div>
            <div className={styles.modalFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                onClick={() => setDialog(EMPTY_DIALOG)}
              />
              <Button
                label={dialog.type === flags.add ? "Create" : "Update"}
                icon={dialog.type === flags.add ? "pi pi-plus" : "pi pi-save"}
                className="submitBtn"
                onClick={sponsorValidation}
              />
            </div>
          </Dialog>

          {/* ── Delete confirmation dialog ── */}
          <Dialog
            visible={dialog.condition && dialog.type === flags.delete}
            style={{ width: "30%" }}
            className="dialogCloseIcon"
            onHide={() => {
              console.log("");
            }}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
          >
            <h3 className="modelHeader">Delete Sponsor</h3>
            <div className={styles.dialogbody}>
              <div className={styles.innerDiv}>
                <i className="pi pi-trash" />
                <span>
                  Are you sure you want to delete {dialog.data.sponsor}?
                </span>
              </div>
            </div>
            <div className={styles.modalFooter}>
              <Button
                className="cancelBtn"
                label="No"
                icon="pi pi-times"
                onClick={() => setDialog(EMPTY_DIALOG)}
              />
              <Button
                label="Yes"
                icon="pi pi-trash"
                className="submitBtn"
                onClick={() => {
                  setLoader(true);
                  updateSponsor(dialog, true);
                }}
              />
            </div>
          </Dialog>
        </div>
      )}
    </>
  );
};

export default Sponsor;
