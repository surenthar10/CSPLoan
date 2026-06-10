import * as React from "react";
import "../assets/CSS/style.css";
import TabNavigation from "./TabNavigation/TabNavigation";
import Loan from "./Loan/Loan";
import BulkUpload from "./BulkUpload/BulkUpload";
import Sponsor from "./Sponsor/Sponsor";

const MainComponent: React.FC = () => {
  const [activeTab, setActiveTab] = React.useState("Loan");
  return (
    <div>
      <div>
        <TabNavigation activeTab={activeTab} onTabChange={setActiveTab} />
        {activeTab === "Loan" && <Loan />}
        {activeTab === "BulkUpload" && <BulkUpload />}
        {activeTab === "Sponsor" && <Sponsor />}
      </div>
    </div>
  );
};

export default MainComponent;
