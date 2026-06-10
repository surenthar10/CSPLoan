import SPServices from "../SPServices/SPServices";
import { errFunc, toastFunc } from "./Config";

export const addItem = async (
  listName: string,
  data: any,
  funcName: string
): Promise<any> => {
  try {
    const _res: any = await SPServices.SPAddItem({
      Listname: listName,
      RequestJSON: data,
    });
    return _res;
  } catch (err) {
    errFunc(err, funcName);
    toastFunc("error", "Error", err);
  }
};
export const updateItem = async (
  listName: string,
  data: any,
  id: number | null,
  funcName: string
): Promise<any> => {
  try {
    const _res: any = await SPServices.SPUpdateItem({
      Listname: listName,
      RequestJSON: data,
      ID: id,
    });
    return _res;
  } catch (err) {
    errFunc(err, funcName);
    toastFunc("error", "Error", err);
  }
};
