import {
  IListItems,
  IAddList,
  IUpdateList,
  ISPGrpMember,
} from "./ISPServicesProps";
import { IItemAddResult, sp } from "@pnp/sp/presets/all";

const SPAddItem = async (params: IAddList): Promise<IItemAddResult> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.add(params.RequestJSON);
};

const SPUpdateItem = async (params: IUpdateList): Promise<IItemAddResult> => {
  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .update(params.RequestJSON);
};
const formatInputs = (data: IListItems): IListItems => {
  if (!data.Select) data.Select = "*";
  if (!data.Topcount) data.Topcount = 5000;
  if (!data.Orderby) data.Orderby = "ID";
  if (!data.Expand) data.Expand = "";
  if (data.Orderbydecorasc !== true && data.Orderbydecorasc !== false)
    data.Orderbydecorasc = true;
  if (!data.PageCount) data.PageCount = 10;
  if (!data.PageNumber) data.PageNumber = 1;

  return data;
};
const formatFilterValue = (params: any[], filterCondition: string): string => {
  let strFilter: string = "";
  if (params) {
    for (let i = 0; i < params.length; i++) {
      if (params[i].FilterKey) {
        if (i != 0) {
          if (filterCondition == "and" || filterCondition == "or") {
            strFilter += " " + filterCondition + " ";
          } else {
            strFilter += " and ";
          }
        }

        if (
          params[i].Operator.toLocaleLowerCase() == "eq" ||
          params[i].Operator.toLocaleLowerCase() == "ne" ||
          params[i].Operator.toLocaleLowerCase() == "gt" ||
          params[i].Operator.toLocaleLowerCase() == "lt" ||
          params[i].Operator.toLocaleLowerCase() == "ge" ||
          params[i].Operator.toLocaleLowerCase() == "le"
        )
          strFilter +=
            params[i].FilterKey +
            " " +
            params[i].Operator +
            "'" +
            params[i].FilterValue +
            "'";
        else if (params[i].Operator.toLocaleLowerCase() == "substringof")
          strFilter +=
            params[i].Operator +
            "('" +
            params[i].FilterValue +
            "','" +
            params[i].FilterKey +
            "')";
      }
    }
  }
  return strFilter;
};
const SPReadItems = async (params: IListItems): Promise<void> => {
  params = formatInputs(params);
  let filterValue: string = formatFilterValue(
    params.Filter ?? [],
    params.FilterCondition ? params.FilterCondition : ""
  );

  return await sp.web.lists
    .getByTitle(params.Listname)
    .items.select(params.Select)
    .filter(filterValue)
    .expand(params.Expand)
    .top(params.Topcount)
    .orderBy(params.Orderby, params.Orderbydecorasc)
    .get();
};
const getSPGroupMember = async (params: ISPGrpMember): Promise<[]> => {
  return await sp.web.siteGroups.getByName(params.GroupName).users.get();
};

export default {
  SPReadItems,
  SPAddItem,
  SPUpdateItem,
  getSPGroupMember,
};
