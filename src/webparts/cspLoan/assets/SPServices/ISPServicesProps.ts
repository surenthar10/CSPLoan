export interface IFilter {
  FilterKey: string;
  FilterValue: string;
  Operator: string;
}
export interface IListItems {
  Listname: string;
  Select?: string | any;
  Topcount?: number | any;
  Expand?: string | any;
  Orderby?: string | any;
  Orderbydecorasc?: boolean;
  Filter?: IFilter[] | any[];
  FilterCondition?: string | any;
  PageCount?: number | any;
  PageNumber?: number | any;
}
export interface IAddList {
  Listname: string;
  RequestJSON: object;
}
export interface IUpdateList {
  Listname: string;
  RequestJSON: object;
  ID: any;
}
export interface ISPGrpMember {
  GroupName: string;
}
