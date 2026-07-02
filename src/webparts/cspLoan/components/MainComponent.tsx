import * as React from "react";
import "../assets/CSS/style.css";
import TabNavigation from "./TabNavigation/TabNavigation";
import Loan from "./Loan/Loan";
import BulkUpload from "./BulkUpload/BulkUpload";
import Sponsor from "./Sponsor/Sponsor";
import { Toast } from "primereact/toast";
import { errFunc, toastRef } from "../assets/Config/Config";
import { useEffect, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { useDispatch } from "react-redux";
import { setIsAdmin } from "../assets/Redux/Features/MainSPContextSlice";
import Loader from "./Loader/Loader";

interface IProps {
  context: any;
}

const MainComponent = (props: IProps) => {
  const dispatch: any = useDispatch();
  const [activeTab, setActiveTab] = React.useState("Loan");
  const toastRefs = toastRef;
  const [loader, setLoader] = useState(false);

  const getOwner = async (_curUser: any) => {
    const users = await sp.web.siteGroups.getById(3).users();
    const isAdmin = users.some((e: any) =>
      e.Email?.toLowerCase().includes(_curUser.email.toLowerCase()),
    );

    dispatch(setIsAdmin(isAdmin));
    setLoader(false);
  };

  const getCurrentUserDetails = async () => {
    await sp.web.currentUser
      .get()
      .then(async (res: any) => {
        let _curUser: any = {
          id: res.Id,
          name: res.Title,
          email: res.Email.toLowerCase(),
        };
        await getOwner(_curUser);
      })
      .catch((err: any) => errFunc(err, "getCurrentUserDetails"));
  };

  const init = async () => {
    setLoader(true);
    await getCurrentUserDetails();
  };

  useEffect(() => {
    void init();
  }, []);

  return (
    <div>
      {loader ? (
        <Loader />
      ) : (
        <>
          <div>
            <TabNavigation activeTab={activeTab} onTabChange={setActiveTab} />

            {activeTab === "Loan" && <Loan context={props.context} />}

            {activeTab === "Bulk Upload" && <BulkUpload />}

            {activeTab === "Sponsor" && <Sponsor />}
          </div>

          <Toast ref={toastRefs} />
        </>
      )}
    </div>
  );
};

export default MainComponent;
