import { IGroup } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReimbursementRequestViewTableProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;

  AdminApprover: string;
  CategoryList: string;
  ClientList: string;
  DepartmentsList: string;
  ProgramList: string;
  ProjectList: string;
  ReimbursementRequestList: string;
  ReimbursementItemsList: string;
  ReimbursementRequestSettingsList: string;
  SubcategoryList: string;
  TasksList: string;
}

export interface IReimbursementRequestViewTableState {
  materialDatas: any;
  // materialRequestData: any;
  groups?: IGroup[];
  expandedItems: any[];
  statusGroups: any[];
  combinedGroups: any[];
  adminApproverId: number;
  departmentName: '';
  HOSName: null;
  Departmentslist: [];
  getcurrentuserId: number;
  noItems: any;
  statusMessageNoItems: string;
  currentPage: number;
  itemsPerPage: number;

}

