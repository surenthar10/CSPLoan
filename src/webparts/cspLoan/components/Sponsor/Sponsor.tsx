/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/no-use-before-define */
import { sp } from "@pnp/sp/presets/all";
import * as React from "react";
import { useEffect, useState } from "react";
import {
  IDialogDetails,
  ILoanRecord,
  ISponsorLoanGroup,
  ISponsorRecord,
} from "../../assets/Config/interface";
import {
  actions,
  flags,
  listNames,
  sponsorUpdateConfig,
  buildSponsorLookupUpdatePayload,
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
import { MultiSelect } from "primereact/multiselect";
import { useSelector } from "react-redux";

// ─── Avatar color palette (one color for all avatars) ─────────────────────────
const AVATAR_COLOR = "#4f46e5";
let updatedLoans: ILoanRecord[] = [];
// ─── Default empty state values ───────────────────────────────────────────────
const EMPTY_SPONSOR: ISponsorRecord = {
  id: null,
  sponsor: "",
  description: "",
  loans: [],
  loanSortLabel: "",
  createdby: "",
  date: null,
  modifieddate: null,
};

const formatSponsorDate = (raw: Date | string | null | undefined): string => {
  if (!raw) return "";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return "";
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const year = date.getFullYear();
  return `${month}/${day}/${year}`;
};

const buildLoanSortLabel = (loans: ILoanRecord[]): string =>
  loans
    .map((loan) => loan.name || "")
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }))
    .join(", ");
const EMPTY_DIALOG: IDialogDetails = {
  condition: false,
  data: EMPTY_SPONSOR,
  type: flags.empty,
};
const Sponsor = () => {
  const _loanDetails: ILoanRecord[] = useSelector(
    (e: any) => e.MainSPContext.loanDetails,
  );
  const [sponsorList, setSponsorList] = useState<ISponsorRecord[]>([]);
  const [sponsorFilterData, setSponsorFilterData] = useState<ISponsorRecord[]>(
    [],
  );
  const [searchText, setSearchText] = useState("");
  const [loader, setLoader] = useState(false);
  const [dialog, setDialog] = useState<IDialogDetails>(EMPTY_DIALOG);
  const [selectedLoans, setselectedLoans] = useState<ILoanRecord[]>([]);

  const closeDialog = (): void => {
    setDialog(EMPTY_DIALOG);
    setselectedLoans([]);
  };

  const openEditDialog = (rowData: ISponsorRecord): void => {
    setselectedLoans([]);
    setDialog({
      condition: true,
      type: flags.edit,
      data: { ...rowData },
    });
  };

  const openAddDialog = (): void => {
    setselectedLoans([]);
    setDialog({
      condition: true,
      type: flags.add,
      data: { ...EMPTY_SPONSOR },
    });
  };

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
      const sponsorArray: ISponsorRecord[] = allRows
        .map((item: any) => {
          const loans = updatedLoans.filter(
            (loan: any) => loan.sponsor?.id === item.Id,
          );
          return {
            id: item.Id,
            sponsor: item.Title,
            description: item.Description || "",
            loans,
            loanSortLabel: buildLoanSortLabel(loans),
            createdby: item.Author?.Title,
            date: item.Created,
            modifieddate: item.Modified,
          };
        })
        .sort((a, b) =>
          (a.sponsor || "").localeCompare(b.sponsor || "", undefined, {
            sensitivity: "base",
          }),
        );
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
    const filtered = sponsorList
      .filter(
        (item) =>
          item.sponsor?.toLowerCase().includes(value.toLowerCase()) ||
          item.description?.toLowerCase().includes(value.toLowerCase()),
      )
      .sort((a, b) =>
        (a.sponsor || "").localeCompare(b.sponsor || "", undefined, {
          sensitivity: "base",
        }),
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
      closeDialog();
      await refreshData("New");
    }
  };
  // Run async tasks with a fixed concurrency limit.
  const runWithConcurrency = async <T,>(
    items: T[],
    concurrency: number,
    worker: (item: T) => Promise<void>,
  ): Promise<void> => {
    if (!items.length) return;

    let nextIndex = 0;
    const executeWorker = async (): Promise<void> => {
      while (nextIndex < items.length) {
        const currentIndex = nextIndex++;
        await worker(items[currentIndex]);
      }
    };

    await Promise.all(
      Array.from({ length: Math.min(concurrency, items.length) }, executeWorker),
    );
  };

  // Fetch all recursive items under a loan folder path.
  const getFolderItems = async (folderPath: string): Promise<any[]> => {
    if (!folderPath) return [];

    const list = sp.web.lists.getByTitle(listNames.loan);
    const allItems: any[] = [];
    let paging: string | undefined;

    do {
      const response: any = await list.renderListDataAsStream({
        ViewXml: `
        <View Scope="RecursiveAll">
          <Query>
            <Where>
              <BeginsWith>
                <FieldRef Name="FileRef" />
                <Value Type="Text">${folderPath}</Value>
              </BeginsWith>
            </Where>
          </Query>
          <ViewFields>
            <FieldRef Name="ID" />
            <FieldRef Name="FileRef" />
            <FieldRef Name="FSObjType" />
          </ViewFields>
          <RowLimit Paged="TRUE">2000</RowLimit>
        </View>
      `,
        Paging: paging,
      });

      allItems.push(...(response.Row || []));

      paging = response.NextHref ? response.NextHref.split("?")[1] : undefined;
    } while (paging);

    return allItems;
  };
  // Collect list item IDs and target sponsor values for all loan groups.
  const collectSponsorAssociationUpdates = async (
    groups: ISponsorLoanGroup[],
  ): Promise<Map<number, number | null>> => {
    const updates = new Map<number, number | null>();

    const fetchTasks: Promise<void>[] = [];

    groups.forEach((group: ISponsorLoanGroup) => {
      group.loans.forEach((loan: ILoanRecord) => {
        fetchTasks.push(
          (async (): Promise<void> => {
            if (!loan.id) return;

            updates.set(loan.id, group.sponsorId);

            const childItems = await getFolderItems(loan.fileRef);
            childItems.forEach((item) => {
              updates.set(Number(item.ID), group.sponsorId);
            });
          })(),
        );
      });
    });

    await Promise.all(fetchTasks);

    return updates;
  };

  // Apply sponsor lookup updates to loan library items in parallel batches.
  const applySponsorAssociationUpdates = async (
    updates: Map<number, number | null>,
  ): Promise<void> => {
    const entries = Array.from(updates.entries());
    if (!entries.length) return;

    const chunks: [number, number | null][][] = [];
    for (
      let i = 0;
      i < entries.length;
      i += sponsorUpdateConfig.batchSize
    ) {
      chunks.push(entries.slice(i, i + sponsorUpdateConfig.batchSize));
    }

    await runWithConcurrency(
      chunks,
      sponsorUpdateConfig.batchConcurrency,
      async (chunk) => {
        const batch = sp.createBatch();

        chunk.forEach(([itemId, sponsorId]) => {
          sp.web.lists
            .getByTitle(listNames.loan)
            .items.getById(itemId)
            .inBatch(batch)
            .update(buildSponsorLookupUpdatePayload(sponsorId));
        });

        await batch.execute();
      },
    );
  };

  // Update sponsor lookup on loan folders and all nested library items.
  const updateSponsorAssociation = async (
    groups: ISponsorLoanGroup[],
  ): Promise<void> => {
    try {
      if (!groups.some((group) => group.loans.length)) return;

      const updates = await collectSponsorAssociationUpdates(groups);
      await applySponsorAssociationUpdates(updates);
    } catch (error) {
      console.error("updateSponsorAssociation error:", error);
      throw error;
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
    const deleteLoans = addData.data.loans.filter(
      (loan: ILoanRecord) => loan.type === "delete",
    );
    const newloans = addData.data.loans.filter(
      (loan: ILoanRecord) => loan.type === "add",
    );

    if (response) {
      const associationGroups: ISponsorLoanGroup[] = [];

      // On sponsor delete, clear Sponsor lookup on every linked loan (reuse empty path).
      if (isDelete) {
        const linkedLoans = updatedLoans.filter(
          (loan: ILoanRecord) =>
            !!loan.id && Number(loan.sponsor?.id) === Number(addData.data.id),
        );

        if (linkedLoans.length > 0) {
          updatedLoans = updatedLoans.map((loan: ILoanRecord) =>
            linkedLoans.some((selected: ILoanRecord) => selected.id === loan.id)
              ? {
                  ...loan,
                  type: "existing",
                  sponsor: {
                    id: 0,
                    sponsorTitle: "",
                  },
                }
              : loan,
          );

          associationGroups.push({ loans: linkedLoans, sponsorId: null });
        }
      } else {
        if (deleteLoans.length > 0) {
          updatedLoans = updatedLoans.map((loan: ILoanRecord) =>
            deleteLoans.some((selected: ILoanRecord) => selected.id === loan.id)
              ? {
                  ...loan,
                  type: "existing",
                  sponsor: {
                    id: 0,
                    sponsorTitle: "",
                  },
                }
              : loan,
          );

          associationGroups.push({ loans: deleteLoans, sponsorId: null });
        }

        if (newloans.length > 0) {
          updatedLoans = updatedLoans.map((loan: ILoanRecord) =>
            newloans.some((selected: ILoanRecord) => selected.id === loan.id)
              ? {
                  ...loan,
                  type: "existing",
                  sponsor: {
                    id: addData.data.id,
                    sponsorTitle: addData.data.sponsor,
                  },
                }
              : loan,
          );

          associationGroups.push({
            loans: newloans,
            sponsorId: addData.data.id,
          });
        }
      }

      if (associationGroups.length > 0) {
        await updateSponsorAssociation(associationGroups);
      }

      closeDialog();
      await refreshData(isDelete ? "Delete" : "Update");
    }
  };
  // ─── Validate form and trigger add or update ──────────────────────────────
  const sponsorValidation = async (): Promise<void> => {
    const sponsorName = dialog.data.sponsor.trim();

    if (!sponsorName) {
      toastFunc("warn", "Warning", "Please enter sponsor");
      return;
    }

    const isDuplicate = sponsorList.some(
      (item: ISponsorRecord) =>
        item.sponsor.toLowerCase() === sponsorName.toLowerCase() &&
        item.id !== dialog.data.id, // Ignore current record during edit
    );

    if (isDuplicate) {
      toastFunc("warn", "Warning", "Sponsor already exists");
      return;
    }

    // Loan selected in dropdown but not added via the Add button.
    if (dialog.type === flags.edit && selectedLoans?.length) {
      toastFunc("warn", "Warning", "Please add the selected loan number");
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
      {rowData?.description || "-"}
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
  const createdOnBodyTemplate = (
    rowData: ISponsorRecord,
  ): React.ReactElement => (
    <span>{formatSponsorDate(rowData.date)}</span>
  );

  const modifiedOnBodyTemplate = (
    rowData: ISponsorRecord,
  ): React.ReactElement => (
    <span>{formatSponsorDate(rowData.modifieddate)}</span>
  );
  //loan number
  const loansBodyTemplate = (rowData: any): React.ReactElement => {
    const loans = rowData.loans || [];
    if (!loans.length) {
      return <span>-</span>;
    }

    return (
      <div className={styles.loanBadges}>
        {loans.map((loanNumber: ILoanRecord) => (
          <span key={loanNumber.id} className={styles.loanBadge}>
            {loanNumber?.name || "-"}
          </span>
        ))}
      </div>
    );
  };
  // ─── Remove a loan number from the associated list ──────────────────────────
  const onDisassociateLoan = (selectedLoan: ILoanRecord): void =>
    setDialog((prev) => ({
      ...prev,
      data: {
        ...prev.data,
        loans: prev.data.loans.map((loan: ILoanRecord) =>
          loan.id === selectedLoan.id ? { ...loan, type: "delete" } : loan,
        ),
      },
    }));
  //add loan value
  const onAddLoan = (): void => {
    if (!selectedLoans?.length) return;
    setDialog((prev) => ({
      ...prev,
      data: {
        ...prev.data,
        loans: [
          ...prev.data.loans,
          ...selectedLoans.map((loan: ILoanRecord) => ({
            ...loan,
            type: "add",
          })),
        ],
      },
    }));

    setselectedLoans([]);
  };
  // ─── Column: edit and delete action icons ─────────────────────────────────
  const actionTemplate = (rowData: ISponsorRecord): React.ReactElement => (
    <div className={styles.actionIcons}>
      <button
        type="button"
        className={`${styles.actionBtn} ${styles.actionBtnEdit}`}
        title="Edit sponsor"
        aria-label={`Edit ${rowData.sponsor || "sponsor"}`}
        onClick={() => openEditDialog(rowData)}
      >
        <i className="pi pi-pencil" />
      </button>
      <button
        type="button"
        className={`${styles.actionBtn} ${styles.actionBtnDelete}`}
        title="Delete sponsor"
        aria-label={`Delete ${rowData.sponsor || "sponsor"}`}
        onClick={() =>
          setDialog({
            ...dialog,
            condition: true,
            type: flags.delete,
            data: { ...rowData },
          })
        }
      >
        <i className="pi pi-trash" />
      </button>
    </div>
  );
  // ─── Load sponsors on mount ───────────────────────────────────────────────
  useEffect(() => {
    setLoader(true);
    getSponsorData("").catch((error) => console.error(error));
  }, []);
  useEffect(() => {
    if (_loanDetails.length) {
      updatedLoans = _loanDetails;
    }
  }, [_loanDetails]);
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
              <i className={`pi pi-search ${styles.searchIcon}`} />
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
              onClick={openAddDialog}
            />
          </div>
          {/* ── Sponsor data table ── */}
          <div className={styles.tableContainer}>
            <DataTable
              value={sponsorFilterData}
              // paginator
              // rows={10}
              // rowsPerPageOptions={[10, 25, 50, 100]}
              emptyMessage="No sponsors found"
              tableStyle={{ width: "100%", tableLayout: "fixed" }}
              paginatorTemplate="CurrentPageReport RowsPerPageDropdown FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink"
              currentPageReportTemplate="Showing {first} to {last} of {totalRecords} records"
              className={styles.sponsorTable}
            >
              <Column
                field="sponsor"
                header="Sponsor"
                sortable
                style={{ width: "14%" }}
              />
              <Column
                field="description"
                body={descriptionTemplate}
                header="Description"
                sortable
                style={{ width: "18%" }}
              />
              <Column
                field="loanSortLabel"
                body={loansBodyTemplate}
                header="Loans"
                sortable
                style={{ width: "22%" }}
              />
              <Column
                field="createdby"
                header="Created By"
                sortable
                style={{ width: "16%" }}
                body={createdByBodyTemplate}
              />
              <Column
                field="date"
                header="Created On"
                sortable
                style={{ width: "11%" }}
                body={createdOnBodyTemplate}
              />
              <Column
                field="modifieddate"
                header="Modified On"
                sortable
                style={{ width: "11%" }}
                body={modifiedOnBodyTemplate}
              />
              <Column
                header="Action"
                style={{ width: "8%" }}
                body={actionTemplate}
              />
            </DataTable>
          </div>
          {/* ── Add / Edit dialog ── */}
          <Dialog
            visible={dialog.condition && dialog.type !== flags.delete}
            className={styles.sponsorDialog}
            style={{
              width: dialog.type === flags.edit ? "560px" : "560px",
            }}
            onHide={closeDialog}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            modal
          >
            <div className={styles.sponsorDialogHeader}>
              <h3 className={styles.sponsorDialogTitle}>
                {dialog.type === flags.add ? actions.add : actions.edit} Sponsor
              </h3>
              {dialog.type === flags.edit && (
                <p className={styles.sponsorDialogSubtitle}>
                  Update details and manage loan associations
                </p>
              )}
            </div>

            <div className={styles.sponsorDialogBody}>
              <div className={styles.sponsorField}>
                <label className={styles.sponsorFieldLabel}>
                  Sponsor <span className={styles.required}>*</span>
                </label>
                <InputText
                  className={`singlelineText ${styles.sponsorInput}`}
                  placeholder="Enter sponsor"
                  value={dialog.data.sponsor}
                  onChange={(e) => onChangeHandler("sponsor", e.target.value)}
                />
              </div>

              <div className={styles.sponsorField}>
                <label className={styles.sponsorFieldLabel}>Description</label>
                <InputTextarea
                  className={`singlelineText ${styles.sponsorTextarea}`}
                  placeholder="Enter description"
                  value={dialog.data.description}
                  onChange={(e) =>
                    onChangeHandler("description", e.target.value)
                  }
                />
              </div>

              {dialog.type === flags.edit && (
                <div className={styles.sponsorField}>
                  <label className={styles.sponsorFieldLabel}>
                    Associated Loans (
                    {
                      (dialog.data.loans || []).filter(
                        (loan: ILoanRecord) => loan.type !== "delete",
                      ).length
                    }
                    )
                  </label>

                  <div className={styles.loansBox}>
                    {(dialog.data.loans || []).filter(
                      (item: ILoanRecord) => item.type !== "delete",
                    ).length === 0 && (
                      <div className={styles.loansEmpty}>
                        No loans linked to this sponsor yet.
                      </div>
                    )}

                    {(dialog.data.loans || [])
                      .filter((item: ILoanRecord) => item.type !== "delete")
                      .map((loan: ILoanRecord) => (
                        <div key={loan.id} className={styles.loanRow}>
                          <span className={styles.loanRowBadge}>
                            {loan.name}
                          </span>
                          <span className={styles.loanRowStatus}>
                            Loan folder linked
                          </span>
                          <button
                            type="button"
                            className={styles.disassociateBtn}
                            onClick={() => onDisassociateLoan(loan)}
                          >
                            <i className="pi pi-times" /> Disassociate
                          </button>
                        </div>
                      ))}

                    <div className={styles.loanAddRow}>
                      <MultiSelect
                        key={`loan-select-${dialog.data.id ?? "new"}`}
                        value={selectedLoans}
                        options={updatedLoans.filter(
                          (loan: ILoanRecord) =>
                            !loan.sponsor || !loan.sponsor.id,
                        )}
                        onChange={(e) => setselectedLoans(e.value)}
                        optionLabel="name"
                        placeholder="Select loan # e.g. 3000115"
                        className={styles.sponsorLoanSelect}
                        filter
                      />
                      <Button
                        label="Add"
                        icon="pi pi-plus"
                        className={styles.sponsorLoanAddBtn}
                        onClick={onAddLoan}
                        disabled={!selectedLoans?.length}
                      />
                    </div>
                  </div>
                </div>
              )}
            </div>

            <div className={styles.sponsorDialogFooter}>
              <Button
                className="cancelBtn"
                icon="pi pi-times"
                label="Cancel"
                onClick={closeDialog}
              />
              <Button
                label={dialog.type === flags.add ? "Create" : "Update"}
                icon={dialog.type === flags.add ? "pi pi-plus" : "pi pi-check"}
                className={styles.sponsorDialogSubmit}
                onClick={sponsorValidation}
              />
            </div>
          </Dialog>

          {/* ── Delete confirmation dialog ── */}
          <Dialog
            visible={dialog.condition && dialog.type === flags.delete}
            className={styles.deleteDialog}
            style={{ width: "440px" }}
            onHide={closeDialog}
            showCloseIcon={false}
            showHeader={false}
            draggable={false}
            modal
          >
            <div className={styles.deleteDialogInner}>
              <div className={styles.deleteDialogIcon}>
                <i className="pi pi-trash" />
              </div>

              <h3 className={styles.deleteDialogTitle}>Delete Sponsor</h3>

              <p className={styles.deleteDialogMessage}>
                Are you sure you want to delete this item?
              </p>

              <div
                className={styles.deleteDialogFileName}
                title={dialog.data.sponsor || "this sponsor"}
              >
                {dialog.data.sponsor || "this sponsor"}
              </div>

              <div className={styles.deleteDialogFooter}>
                <Button
                  className="cancelBtn"
                  label="Cancel"
                  icon="pi pi-times"
                  iconPos="left"
                  onClick={closeDialog}
                />
                <Button
                  label="Delete"
                  icon="pi pi-trash"
                  iconPos="left"
                  className={styles.deleteDialogConfirm}
                  onClick={() => {
                    setLoader(true);
                    updateSponsor(dialog, true);
                  }}
                />
              </div>
            </div>
          </Dialog>
        </div>
      )}
    </>
  );
};

export default Sponsor;
