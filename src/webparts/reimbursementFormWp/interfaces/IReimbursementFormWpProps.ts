import { IDropdownOption } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReimbursementFormWpProps {
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

export interface IReimbursementFormWpState {
  currentDate: string;
  listClient: any[];
  client: any;
  selectedClient: IDropdownOption;
  selectedCategory: IDropdownOption;
  selectedProgram: IDropdownOption;
  getProject: IDropdownOption;
  getSubcategory: IDropdownOption;
  listProgram: any[];
  listSubcategory: any[];
  listProject: any[];
  listCategory: any[];
  program: any;
  project: any;
  subcategory: any;
  category: any;
  department: string;
  expenseDetails: string;
  amount: string; 
  rows: any;
  datepicker : any;
  files: any;
  totalAmount: Number;
  advanceAmount: Number;
  balanceAmount: Number;
  comment: string;
  purpose: string;
  HOSName: Number;
  Departmentslist: any;
  departmentName: string;
  ReimbursementRequestID: number;
  taskListItemId: number;
  isPopupVisible: boolean;
  isOkButtonDisabled: boolean;
  displayName: string;
  designation: string;
  adminApproverName: string;
  HOSNameStringId: string;
  HOSNameEmail: string;
  adminApproverEmail: string;
  employeeID: Number;

}

export interface IReimbursementFormWpWebPartProps {
  description: string;
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