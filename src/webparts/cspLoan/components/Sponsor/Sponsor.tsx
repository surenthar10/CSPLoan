/*@typescript-eslint/no-floating-promises*/
import { sp } from "@pnp/sp/presets/all";
import * as React from "react";
import { useEffect } from "react";
import { ISponsorRecord } from "../../assets/Config/interface";
import { listNames } from "../../assets/Config/Config";
import { InputText } from "primereact/inputtext";
import { Button } from "primereact/button";
import style from "./Sponsor.module.scss";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
const Sponsor = () => {
  // State to hold sponsor list
  const [sponsorList, setSponsorList] = React.useState<ISponsorRecord[]>([]);
  //Sposor List
  const getSponsorData = async () => {
    try {
      let paged = await sp.web.lists
        .getByTitle(listNames.sponsors)
        .items.select("*", "Author/Title")
        .expand("Author")
        .top(5000)
        .getPaged();
      const allRows = [...paged.results];
      while (paged.hasNext) {
        paged = await paged.getNext();
        allRows.push(...paged.results);
      }
      console.log(allRows);
      setSponsorList(allRows);
    } catch (error) {
      console.error(error);
    }
  };
  //fist last letter
  function getInitials(name: string): string {
    if (!name) return "?";
    const parts = name.trim().split(" ");
    return parts.length >= 2
      ? `${parts[0][0]}${parts[parts.length - 1][0]}`.toUpperCase()
      : name.slice(0, 2).toUpperCase();
  }
  // ─── Created By column: avatar + name ────────────────────────────────────────
  const createdByBodyTemplate = (rowData: any): React.ReactElement => {
    const name: string = rowData.AuthorName ?? rowData.Author?.Title ?? "";
    return (
      <div className={style.createdByCell}>
        <span className={style.avatar}>{getInitials(name)}</span>
        <span className={style.createdByName}>{name}</span>
      </div>
    );
  };
  // ─── Created On column: DD/MM/YYYY ───────────────────────────────────────────
  const createdOnBodyTemplate = (rowData: any): React.ReactElement => {
    const raw = rowData.Created ?? rowData.Modified ?? "";
    let display = "";
    if (raw) {
      const d = new Date(raw);
      const dd = String(d.getDate()).padStart(2, "0");
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const yyyy = d.getFullYear();
      display = `${dd}/${mm}/${yyyy}`;
    }
    return <span>{display}</span>;
  };
  //description
  const descriptionTemplate = (rowData: any) => {
    return (
      <div
        title={rowData.Description}
        style={{
          maxWidth: "350px",
          overflow: "hidden",
          whiteSpace: "nowrap",
          textOverflow: "ellipsis",
          cursor: "pointer",
        }}
      >
        {rowData.Description}
      </div>
    );
  };
  //Action template
  const actionTemplate = (rowData: any) => {
    return (
      <div className={style.actionIcons}>
        <i className="pi pi-pencil" />

        <i className="pi pi-trash" />
      </div>
    );
  };
  // useEffect to get sponsor data on component mount
  useEffect(() => {
    getSponsorData().catch((error) => console.error(error));
  }, []);

  return (
    <div className={style.sponsorContainer}>
      {/* // Toolbar with title, search, refresh and new sponsor button */}
      <div className={style.sponsortoolbar}>
        <div className={style.toolbarTitle}>Sponsor Details</div>

        <div className={style.toolbarSearch}>
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
          ></i>
          <InputText
            placeholder="Search sponsors..."
            className={style.searchInput}
          />
        </div>
        <Button icon="pi pi-refresh" className={style.refreshBtn} />
        <Button
          label="New Sponsor"
          icon="pi pi-plus"
          className={style.newSponsorBtn}
        />
      </div>
      {/* // Table to display sponsor data */}
      <div className={style.tableContainer}>
        <DataTable
          value={sponsorList}
          paginator
          rows={10}
          rowsPerPageOptions={[10, 25, 50, 100]}
          emptyMessage="No sponsors found"
          tableStyle={{ minWidth: "100%" }}
          paginatorTemplate="CurrentPageReport RowsPerPageDropdown FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink"
          currentPageReportTemplate="Showing {first} to {last} of {totalRecords} records"
          className={style.sponsorTable}
        >
          <Column
            field="Title"
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
            field="Created"
            header="Created On"
            sortable
            body={createdOnBodyTemplate}
            style={{ width: "12%" }}
          />
          <Column
            header="Action"
            style={{ width: "9%" }}
            body={actionTemplate}
          />
        </DataTable>
      </div>
    </div>
  );
};

export default Sponsor;
