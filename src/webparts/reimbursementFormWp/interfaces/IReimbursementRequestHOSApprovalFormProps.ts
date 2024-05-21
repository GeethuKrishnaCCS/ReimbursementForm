import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReimbursementRequestHOSApprovalFormProps {
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

export interface IReimbursementRequestHOSApprovalFormState {
  getItemId: string;
  RequestedBy: string;
  getcurrentuserId: number;
  reimbursementRequestData: any;
  RequestedDate: string;
  client: string;
  program: string;
  project: string;
  department: string;
  advanceAmount: string;
  balanceAmount: string;
  totalAmount: string;
  reimbursementRequestDataId: number;
  reimbursementListData: any;
  RequestorComments: string;
  comment: string;
  FinanceApproverId: Number;
  taskListItemId: Number;
  HOSApproverId: Number;
  RequestedById: Number;
  isOkButtonDisabled: boolean;
  isTaskIdPresent: any;
  noAccessId: any;
  statusMessageTAskIdNull: string;
}

