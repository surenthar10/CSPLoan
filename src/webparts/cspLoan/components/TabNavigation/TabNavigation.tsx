import * as React from "react";
import styles from "./TabNavigation.module.scss";
import { useSelector } from "react-redux";

const TabNavigation = (props: {
  activeTab: string;
  onTabChange: (tab: string) => void;
}) => {
  const isAdmin: boolean = useSelector((e: any) => e.MainSPContext.isAdmin);
  let tabs = [];

  if (isAdmin) {
    tabs = ["Loan", "Bulk Upload", "Sponsor"];
  } else {
    tabs = ["Loan", "Bulk Upload"];
  }

  return (
    <div>
      <>
        <div className={styles.headerContainer}>
          {/* Logo Section */}
          <div className={styles.logoSection}>
            <img
              src={require("../../assets/images/cpclogo.svg")}
              alt="CPC Logo"
              className={styles.logo}
            />

            <div className={styles.companyText}>
              <div>Community</div>
              <div>Preservation</div>
              <div>Corporation</div>
            </div>
          </div>

          {/* Navigation Tabs */}
          <div className={styles.navSection}>
            {tabs.map((tab) => (
              <button
                key={tab}
                className={
                  props.activeTab === tab
                    ? `${styles.tabButton} ${styles.active}`
                    : styles.tabButton
                }
                onClick={() => props.onTabChange(tab)}
              >
                {tab}
              </button>
            ))}
          </div>
        </div>
      </>
    </div>
  );
};

export default TabNavigation;
