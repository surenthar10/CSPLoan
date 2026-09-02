import * as React from "react";
import { useMemo } from "react";
import { useSelector } from "react-redux";
import styles from "./TabNavigation.module.scss";

interface ITabItem {
  id: string;
  label: string;
  icon: string;
}

const ALL_TABS: ITabItem[] = [
  { id: "Loan", label: "Loan Files", icon: "pi-folder" },
  { id: "Sponsor", label: "Sponsors", icon: "pi-users" },
  { id: "Bulk Upload", label: "Bulk Upload", icon: "pi-cloud-upload" },
];

const TabNavigation = (props: {
  activeTab: string;
  onTabChange: (tab: string) => void;
}): React.ReactElement => {
  const isAdmin: boolean = useSelector((e: any) => e.MainSPContext.isAdmin);

  const tabs = useMemo(
    () =>
      isAdmin ? ALL_TABS : ALL_TABS.filter((tab) => tab.id !== "Sponsor"),
    [isAdmin],
  );

  return (
    <header className={styles.headerContainer}>
      <nav className={styles.navSection} aria-label="Main navigation">
        {tabs.map((tab) => {
          const isActive = props.activeTab === tab.id;

          return (
            <button
              key={tab.id}
              type="button"
              className={`${styles.tabButton} ${
                isActive ? styles.active : ""
              }`}
              aria-current={isActive ? "page" : undefined}
              onClick={() => props.onTabChange(tab.id)}
            >
              <span className={styles.tabIconWrap} aria-hidden>
                <i className={`pi ${tab.icon} ${styles.tabIcon}`} />
              </span>
              <span className={styles.tabCopy}>
                <span className={styles.tabLabel}>{tab.label}</span>
                {isActive && (
                  <span className={styles.tabHint}>Currently viewing</span>
                )}
              </span>
            </button>
          );
        })}
      </nav>
    </header>
  );
};

export default TabNavigation;
