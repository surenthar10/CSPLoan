import * as React from "react";
import "../assets/CSS/style.css";
import TabNavigation from "./TabNavigation/TabNavigation";
import Loan from "./Loan/Loan";
import BulkUpload from "./BulkUpload/BulkUpload";
import Sponsor from "./Sponsor/Sponsor";
import { Toast } from "primereact/toast";
import { toastRef } from "../assets/Config/Config";

interface IProps {
  context: any;
}

const MainComponent = (props: IProps) => {
  const [activeTab, setActiveTab] = React.useState("Loan");
  const toastRefs = toastRef;
  return (
    <div>
      <div>
        <TabNavigation activeTab={activeTab} onTabChange={setActiveTab} />
        {activeTab === "Loan" && <Loan context={props.context} />}
        {activeTab === "Bulk Upload" && <BulkUpload />}
        {activeTab === "Sponsor" && <Sponsor />}
      </div>
      <Toast ref={toastRefs} />
    </div>
  );
};

export default MainComponent;
