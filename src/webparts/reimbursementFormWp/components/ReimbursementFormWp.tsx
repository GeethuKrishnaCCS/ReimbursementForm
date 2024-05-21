import * as React from 'react';
import styles from './ReimbursementFormWp.module.scss';
import { IReimbursementFormWpProps, IReimbursementFormWpState } from '../interfaces/IReimbursementFormWpProps';
import { DatePicker, DefaultButton, Dropdown, FocusTrapZone, IDropdownOption, IIconProps, IconButton, Label, Layer, Overlay, Popup, PrimaryButton, TextField } from '@fluentui/react';
import { ReimbursementFormWpService } from '../services';
import * as moment from 'moment';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import replaceString from 'replace-string';
import ReimbursementRequestHOSApprovalForm from './ReimbursementRequestHOSApprovalForm';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import ReimbursementRequestAdminApprovalForm from './ReimbursementRequestAdminApprovalForm';
import ReimbursementRequestViewTable from './ReimbursementRequestViewTable';
import Toast from './Toast';

export default class ReimbursementFormWp extends React.Component<IReimbursementFormWpProps, IReimbursementFormWpState, {}> {
  private _service: any;
  private fileInput: HTMLInputElement | null = null;


  public constructor(props: IReimbursementFormWpProps) {
    super(props);
    this._service = new ReimbursementFormWpService(this.props.context);

    this.state = {

      currentDate: moment(new Date()).format("DD-MM-YYYY"),
      listClient: [],
      client: [],
      selectedClient: { key: "", text: "" },
      selectedCategory: { key: "", text: "" },
      selectedProgram: { key: "", text: "" },
      getProject: { key: "", text: "" },
      getSubcategory: { key: "", text: "" },
      listProgram: [],
      listSubcategory: [],
      listProject: [],
      listCategory: [],
      program: [],
      subcategory: [],
      project: [],
      category: [],
      department: '',
      expenseDetails: '',
      amount: '',
      // rows: [],
      datepicker: null,
      files: [],
      // rows: [
      //   { datepicker: null, category: "", particulars: '', amount: '', files: [] } // Add initial row here
      // ],
      rows: [
        {
          datepicker: null,
          expenseDetails: '',
          amount: '',
          selectedCategory: { key: "", text: "" }, // Include selectedCategory in the initial row
          getSubcategory: { key: "", text: "" }, // Include getSubcategory in the initial row
          files: []
        }
      ],
      totalAmount: 0,
      advanceAmount: 0,
      balanceAmount: 0,
      comment: "",
      purpose: "",
      HOSName: null,
      Departmentslist: [],
      departmentName: '',
      ReimbursementRequestID: null,
      taskListItemId: null,
      isPopupVisible: false,
      isOkButtonDisabled: false,
      displayName: '',
      designation: '',
      adminApproverName: '',
      HOSNameStringId: "",
      HOSNameEmail: "",
      adminApproverEmail: "",
      employeeID: null,


    }
    // this.getCurrentDate = this.getCurrentDate.bind(this);
    // this.getClientList = this.getClientList.bind(this);
    // this.getProgramList = this.getProgramList.bind(this);
    this.getProjectList = this.getProjectList.bind(this);
    this.getProjectChange = this.getProjectChange.bind(this);
    // this.UserProfiles = this.UserProfiles.bind(this);
    // this.onChangeAmount = this.onChangeAmount.bind(this);
    this.onChangeParticulars = this.onChangeParticulars.bind(this);
    this.OnDateChange = this.OnDateChange.bind(this);
    this.onFormatDate = this.onFormatDate.bind(this);
    this.addRow = this.addRow.bind(this);
    this.deleteRow = this.deleteRow.bind(this);
    this.calculateTotalAmount = this.calculateTotalAmount.bind(this);
    this.onChangeAdvanceAmount = this.onChangeAdvanceAmount.bind(this);
    this.onChangeComment = this.onChangeComment.bind(this);
    this.onChangePurpose = this.onChangePurpose.bind(this);
    this.onSubmitClick = this.onSubmitClick.bind(this);
    this.getDepartmentsList = this.getDepartmentsList.bind(this);
    this.sendEmailNotificationToHOS = this.sendEmailNotificationToHOS.bind(this);
    this.getCurrentUser = this.getCurrentUser.bind(this);
    this.hidePopup = this.hidePopup.bind(this);
    this.onPopOk = this.onPopOk.bind(this);
    this.getCategoryist = this.getCategoryist.bind(this);
    // this.getSubcategoryChange = this.getSubcategoryChange.bind(this);
    this.getSubcategoryList = this.getSubcategoryList.bind(this);
    this.onCategoryChange = this.onCategoryChange.bind(this);
    // this.getHOS = this.getHOS.bind(this);
    this.onClickCancel = this.onClickCancel.bind(this);
    this.getEmployeeCode = this.getEmployeeCode.bind(this);
    this.areRowsComplete = this.areRowsComplete.bind(this);
    this.handleChangeAmount = this.handleChangeAmount.bind(this);



  }


  public async componentDidMount() {
    await this.getCurrentUser();
    await this.UserProfiles();
    await this.getDepartmentsList();
    await this.getHOS();
    await this.getAdminApprover();
    // await this.getCurrentDate();
    await this.getProjectList();
    await this.getCategoryist();
    // await this.getClientList();
    await this.getEmployeeCode();

    const url = this.props.context.pageContext.web.absoluteUrl + "/SitePages/ViewSubmittedRequests.aspx" + "?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js";
    console.log('url: ', url);

  }

  public getCurrentDate() {
    const date = moment(new Date).format("DD-MM-YYYY");
    this.setState({ currentDate: date })
  }
  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    console.log('getcurrentuser: ', getcurrentuser);
  }

  public async getEmployeeCode() {
    const graphClient = await this.props.context.msGraphClientFactory.getClient('3');

    const response = await graphClient.api("users")
      .version("v1.0")
      .select("department,employeeId,displayName,mail")
      .filter(`mail eq '${this.props.context.pageContext.user.email}'`)
      .get();
    // console.log('response: ', response);

    this.setState({
      employeeID: response.value[0].employeeId
    });
    // console.log('employeeID: ', this.state.employeeID);
  }

  public async getAdminApprover() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const AdminApproverlistItem = await this._service.getListItems(this.props.AdminApprover, url)
    const ApproverIdUserInfo = await this._service.getUser(AdminApproverlistItem[0].AdminApproverId);
    this.setState({ adminApproverName: ApproverIdUserInfo.Id });
    this.setState({ adminApproverEmail: ApproverIdUserInfo.Email });
  }

  public async getHOS() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const DepartmentslistItem = await this._service.getClientListItems(this.props.DepartmentsList, url)
    const filteredDepartment = await DepartmentslistItem.find((dept: any) => dept.Title === this.state.department);
    const HOSNameUserInfo = await this._service.getUser(filteredDepartment.HOSNameId);
    const HOSName = HOSNameUserInfo.Email;
    this.setState({ HOSNameEmail: HOSName });
  }


  public async UserProfiles() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl
    const getUsers: string = url + `/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
    await fetch(getUsers, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json'
      },
      credentials: 'include'
    })
      .then(response => response.json())
      .then(async data => {
        // console.log('data: ', data);
        // console.log('department ', data.UserProfileProperties.filter((p: any) => p.Key === 'Department')[0].Value);
        this.setState({
          displayName: data.DisplayName,
          department: await data.UserProfileProperties.filter((p: any) => p.Key === 'Department')[0].Value,
          designation: data.UserProfileProperties.filter((p: any) => p.Key === 'SPS-JobTitle')[0].Value,
        });

      })
      .catch(error => {
        console.error('Error:', error);
      });
  }

  // public async getClientList() {
  //   const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const listClient = await this._service.getClientListItems(this.props.ClientList, url)
  //   this.setState({ listClient: listClient })

  //   const ClientList: any[] = [];
  //   listClient.forEach((client: any) => {
  //     ClientList.push({ key: client.ID, text: client.Client });
  //   });
  //   this.setState({ client: ClientList });
  // }

  // public getProgramList = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => {
  //   this.setState({ selectedClient: item });
  //   const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const listProgram = await this._service.getProgramListItems(this.props.ProgramList, item.key, url)
  //   this.setState({ listProgram: listProgram })

  //   const Program: any[] = [];
  //   listProgram.forEach((programItem: any) => {
  //     Program.push({ key: programItem.ID, text: programItem.Program });
  //   });
  //   this.setState({ program: Program });
  // }

  // public getProjectList = async (event: React.FormEvent<HTMLDivElement>, data: IDropdownOption) => {
  //   this.setState({ selectedProgram: data });

  //   const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const listProject = await this._service.getProjectListItems(this.props.ProjectList, data.key, url)
  //   this.setState({ listProject: listProject })

  //   const ProjectList: any[] = [];
  //   listProject.forEach((project: any) => {
  //     ProjectList.push({ key: project.ID, text: project.Project });
  //   });
  //   this.setState({ project: ProjectList });
  // }

  public async getProjectList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const listProject = await this._service.getListItems(this.props.ProjectList, url)
    this.setState({ listProject: listProject })

    const ProjectList: any[] = [];
    listProject.forEach((project: any) => {
      ProjectList.push({ key: project.ID, text: project.Project });
    });
    this.setState({ project: ProjectList });
  }

  public async getCategoryist() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const listCategory = await this._service.getListItems(this.props.CategoryList, url)
    this.setState({ listCategory: listCategory })

    const categoryList: any[] = [];
    listCategory.forEach((category: any) => {
      categoryList.push({ key: category.ID, text: category.Category });
    });
    this.setState({ category: categoryList });
  }

  public onCategoryChange = async (event: React.FormEvent<HTMLDivElement>, category: IDropdownOption, index: number) => {
    const updatedRows = [...this.state.rows];
    updatedRows[index].selectedCategory = category;
    this.setState({ rows: updatedRows });

    await this.getSubcategoryList(event, category, index);
  };

  // public getSubcategoryList = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => {
  //   this.setState({ selectedCategory: item });

  //   const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const listSubcategory = await this._service.getSubcategoryListItems("SubcategoryList", item.key, url)
  //   console.log('listSubcategory: ', listSubcategory);
  //   this.setState({ listSubcategory: listSubcategory })

  //   const subCategory: any[] = [];
  //   listSubcategory.forEach((subcategory: any) => {
  //     subCategory.push({ key: subcategory.ID, text: subcategory.Subcategory });
  //   });
  //   this.setState({ subcategory: subCategory });
  // }

  public getSubcategoryList = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption, index: number) => {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const listSubcategory = await this._service.getSubcategoryListItems(this.props.SubcategoryList, item.key, url);
    const subCategoryOptions = listSubcategory.map((subcategory: any) => ({ key: subcategory.ID, text: subcategory.Subcategory }));

    const updatedRows = [...this.state.rows];
    updatedRows[index].selectedCategory = item; // Update selected category for the current row
    updatedRows[index].subcategory = subCategoryOptions; // Update subcategory options for the current row
    this.setState({ rows: updatedRows });
  };

  // public getSubcategoryChange(event: React.FormEvent<HTMLDivElement>, getSubcategory: IDropdownOption) {
  //   this.setState({ getSubcategory: getSubcategory });
  // }

  public getProjectChange(event: React.FormEvent<HTMLDivElement>, getProject: IDropdownOption) {
    this.setState({ getProject: getProject });
  }

  public async getDepartmentsList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const DepartmentslistItem = await this._service.getClientListItems(this.props.DepartmentsList, url)
    // console.log('DepartmentslistItem: ', DepartmentslistItem);
    this.setState({ Departmentslist: DepartmentslistItem });

    DepartmentslistItem.map((Item: any) => {
      const departmentName = Item.Title;
      const HOSName = Item.HOSNameId;
      const HOSNameStringId = Item.HOSNameStringId;


      this.setState({
        departmentName: departmentName,
        HOSName: HOSName,
        HOSNameStringId: HOSNameStringId,
      });
    })
  }


  public onChangeParticulars(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, Particulars: string) {
    this.setState({ expenseDetails: Particulars });
  }
  public OnDateChange = (date: Date | null | undefined): void => {
    this.setState({ datepicker: date });
  };

  public onFormatDate = (date: Date): string => {
    let selectd = moment(date).format('DD.MM.YYYY');
    return selectd;
  };

  // public onChangeAmount(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, Amount: string, index: number) {
  //   const updatedRows = [...this.state.rows];
  //   updatedRows[index].amount = Amount;
  //   this.setState({ rows: updatedRows }, () => {
  //     this.calculateTotalAmount(); // Update total amount whenever amount changes
  //   });
  // }

  public handleChangeAmount = (event: any, index: any) => {
    const { value } = event.target;
    const updatedRows = [...this.state.rows];
    // Regular expression to allow only numbers with up to 2 decimal places
    const regex = /^\d*\.?\d{0,2}$/;
    if (value === '' || regex.test(value)) {
      updatedRows[index].amount = value;
      this.setState({ rows: updatedRows }, () => {
        this.calculateTotalAmount();
      });
    }
  };

  public onFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      const fileList: any[] = [];
      for (let i = 0; i < files.length; i++) {
        fileList.push(files[i]);
      }
      this.setState((prevState) => ({
        files: [...prevState.files, ...fileList],
      }), () => {
        console.log('files: ', this.state.files);
      });
    }
  };
  public onChangeComment(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, Comment: string) {
    this.setState({ comment: Comment });
  }

  public onChangePurpose(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, Purpose: string) {
    this.setState({ purpose: Purpose });
  }

  public deleteRow = (index: number) => {
    const updatedRows = [...this.state.rows];
    updatedRows.splice(index, 1);
    this.setState({ rows: updatedRows }, () => {
      this.calculateTotalAmount();
    });
  };

  public addRow = () => {
    const { datepicker, expenseDetails, amount, selectedCategory, getSubcategory } = this.state;
    const newRow = { datepicker, expenseDetails, amount, selectedCategory, getSubcategory };
    const rows = [...this.state.rows, newRow];

    this.setState({ rows }, () => {
      this.calculateTotalAmount();
    });
  };


  public onChangeAdvanceAmount(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, advanceAmount: string) {
    this.setState({ advanceAmount: parseFloat(advanceAmount) || 0 }, () => {
      this.setState({ balanceAmount: this.calculateBalanceAmount() });
    });
  }

  public calculateBalanceAmount() {
    const totalAmount = Number(this.state.totalAmount) || 0;
    const advanceAmount = Number(this.state.advanceAmount) || 0;
    return totalAmount - advanceAmount;
  }

  private calculateTotalAmount() {
    let totalAmount = 0;
    this.state.rows.forEach((row: any) => {
      totalAmount += parseFloat(row.amount) || 0;
    });
    this.setState({ totalAmount }, () => {
      this.setState({ balanceAmount: this.calculateBalanceAmount() });
    });
  }

  public async onSubmitClick(): Promise<void> {
    this.setState({
      isPopupVisible: true,
    });
  }

  public async onPopOk(): Promise<void> {
    await this.setState({ isOkButtonDisabled: true });
    const filteredDepartment = this.state.Departmentslist.find((dept: any) => dept.Title === this.state.department);

    // console.log('currentDate:', this.state.currentDate);

    if (filteredDepartment) {
      const HOSName = filteredDepartment.HOSNameId;

      const dataItem = {
        RequestedDate: moment(new Date()).format("DD-MM-YYYY"),
        // ClientId: this.state.selectedClient.key,
        // ProgramId: this.state.selectedProgram.key,
        // ProjectId: this.state.getProject.key,
        ProjectId: this.state.getProject.key,
        RequestorComments: this.state.comment,
        Department: this.state.department,
        Purpose: this.state.purpose,
        TotalAmountNumber: this.state.totalAmount,
        BalanceAmountNumber: this.state.balanceAmount,
        AdvanceAmountNumber: this.state.advanceAmount,
        HOSApproverId: HOSName,
        // HOSApprovalStatus: "HOS Pending",
        Status: "HOS Pending",
        // HOSApprovalComments
        // HOSApprovedDate
        FinanceApproverId: this.state.adminApproverName,
        // FinanceApproverComments
        // FinanceApprovalStatus
        // FinanceApproverDate
        Designation: this.state.designation,
        EmployeeID: this.state.employeeID
      };

      const url: string = this.props.context.pageContext.web.serverRelativeUrl;
      await this._service.addItemRequestForm(dataItem, this.props.ReimbursementRequestList, url).then(async (item: any) => {
        // console.log('item: ', item);

        const itemId = item.data.Id;
        this.setState({ ReimbursementRequestID: itemId });

        const dataItem = {
          ReimbursementRequestCode: "00" + itemId
        }
        await this._service.updateRequestForm(this.props.ReimbursementRequestList, dataItem, itemId, url);

        const ReimbursementItemsData = this.state.rows.map((row: any) => ({

          ReimbursementRequestIDId: itemId,
          // Date : ,
          ExpenseDetails: row.expenseDetails,
          Amount: row.amount,
          CategoryId: row.selectedCategory.key,
          SubcategoryId: row.getSubcategory.key,
          ExpenseDate: moment(row.datepicker).format("DD-MM-YYYY")
        }));

        const attachment = this.state.files;
        // console.log('attachment: ', attachment);
        for (const ReimbursementItemData of ReimbursementItemsData) {
          await this._service.addItemRequestForm(ReimbursementItemData, this.props.ReimbursementItemsList, url).then(async (Dataitem: any) => {
            // console.log('Dataitem: ', Dataitem);

            if (attachment && attachment.length > 0) {
              for (const file of attachment) {
                await this._service.addAttachments(this.props.ReimbursementItemsList, Dataitem.data.Id, file.name, file);
              }
            }
          })
        }

        const taskItem = {
          ReimbursementRequestIDId: itemId,
          AssignedToId: HOSName,

        }
        await this._service.addItemRequestForm(taskItem, this.props.TasksList, url).then(async (task: any) => {
          // console.log('task: ', task);

          this.setState({ taskListItemId: task.data.Id });
          const taskURL = url + "/SitePages/" + "ReimbursementRequestHOSApprovalForm" + ".aspx?did=" + item.data.Id + "&itemid=" + task.data.Id + "&formType=HOSApproval";
          const taskItemtoupdate = {
            TaskTitleWithLink: {
              Description: "-- HOS Approval",
              Url: taskURL,
            }
          }
          await this._service.updateRequestForm(this.props.TasksList, taskItemtoupdate, task.data.Id, url);
          await this.sendEmailNotificationToHOS(this.props.context);
          this.hidePopup();
          Toast("success", "Successfully Submitted");
          alert("submit")
          setTimeout(() => {
            window.location.href = url;
          }, 3000);
        })
      })


    } else {
      console.error('Department not found');
    }
  }

  public async sendEmailNotificationToHOS(context: any): Promise<void> {
    const filteredDepartment = this.state.Departmentslist.find((dept: any) => dept.Title === this.state.department);

    if (filteredDepartment && filteredDepartment.HOSNameId) {
      const hosApproverId = filteredDepartment.HOSNameId;
      const hosApproverIdUserInfo = await this._service.getUser(hosApproverId);
      const HOSApproverEmail = hosApproverIdUserInfo.Email;
      const HOSApprover = hosApproverIdUserInfo.Title;
      const url: string = this.props.context.pageContext.web.absoluteUrl;
      const evaluationURL = url + "/SitePages/" + "MaterialRequestApprovalForm" + ".aspx?did=" + this.state.ReimbursementRequestID + "&itemid=" + this.state.taskListItemId + "&formType=HOSApproval";

      const project = this.state.project.find((proj: any) => proj.key === this.state.getProject.key)?.text || 'Project';

      const getcurrentuser = await this._service.getCurrentUser();
      const getcurrentUserInfo = await this._service.getUser(getcurrentuser.Id);
      const employeeName = getcurrentUserInfo.Title;

      const requestedDate = this.state.currentDate;

      const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
      const emailNoficationSettings = await this._service.getListItems(this.props.ReimbursementRequestSettingsList, serverurl);
      const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendEmailNotificationToHOS");

      if (emailNotificationSetting) {
        const subjectTemplate = emailNotificationSetting.Subject;
        const bodyTemplate = emailNotificationSetting.Body;

        const replaceSubject = replaceString(subjectTemplate, '[Project]', project)


        const replaceHOSApprover = replaceString(bodyTemplate, '[HOSApprover]', HOSApprover)
        const replaceEmployyeName = replaceString(replaceHOSApprover, '[EmployeeName]', employeeName)
        const replaceProject = replaceString(replaceEmployyeName, '[Project]', project)
        const replaceRequestedDate = replaceString(replaceProject, '[RequestedDate]', requestedDate)
        const replacedBodyWithLink = replaceString(replaceRequestedDate, '[Link]', `<a href="${evaluationURL}" target="_blank">Click here</a>`);

        const emailPostBody: any = {
          message: {
            subject: replaceSubject,
            body: {
              contentType: 'HTML',
              content: replacedBodyWithLink
            },
            toRecipients: [
              {
                emailAddress: {
                  address: HOSApproverEmail,
                },
              },
            ],
          },
        };

        return context.msGraphClientFactory
          .getClient('3')
          .then((client: MSGraphClientV3): void => {
            client.api('/me/sendMail').post(emailPostBody);
          });
      }
    }
  }

  public hidePopup = () => {
    this.setState({ isPopupVisible: false });
  };

  public onClickCancel() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    window.location.href = url;
  };

  public areRowsComplete() {
    return this.state.rows.every((row: any) => {
      return row.selectedCategory && row.getSubcategory && row.amount;
    });
  }


  public render(): React.ReactElement<IReimbursementFormWpProps> {
    const currentUrl = window.location.href;

    // if (currentUrl === this.props.context.pageContext.web.absoluteUrl + "/SitePages/ViewSubmittedRequests.aspx" + "?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js") {
      if (currentUrl === this.props.context.pageContext.web.absoluteUrl + "/SitePages/ViewSubmittedRequests.aspx") {
      return <ReimbursementRequestViewTable
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
        AdminApprover={this.props.AdminApprover}
        CategoryList={this.props.CategoryList}
        ClientList={this.props.ClientList}
        DepartmentsList={this.props.DepartmentsList}
        ProgramList={this.props.ProgramList}
        ProjectList={this.props.ProjectList}
        ReimbursementRequestList={this.props.ReimbursementRequestList}
        ReimbursementItemsList={this.props.ReimbursementItemsList}
        ReimbursementRequestSettingsList={this.props.ReimbursementRequestSettingsList}
        SubcategoryList={this.props.SubcategoryList}
        TasksList={this.props.TasksList}
      />;
    }

    else if (new URLSearchParams(window.location.search).get("formType") === "HOSApproval") {
      return <ReimbursementRequestHOSApprovalForm
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
        AdminApprover={this.props.AdminApprover}
        CategoryList={this.props.CategoryList}
        ClientList={this.props.ClientList}
        DepartmentsList={this.props.DepartmentsList}
        ProgramList={this.props.ProgramList}
        ProjectList={this.props.ProjectList}
        ReimbursementRequestList={this.props.ReimbursementRequestList}
        ReimbursementItemsList={this.props.ReimbursementItemsList}
        ReimbursementRequestSettingsList={this.props.ReimbursementRequestSettingsList}
        SubcategoryList={this.props.SubcategoryList}
        TasksList={this.props.TasksList}
      />;
    }
    else if (new URLSearchParams(window.location.search).get("formType") === "AdminApproval") {
      return <ReimbursementRequestAdminApprovalForm
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
        AdminApprover={this.props.AdminApprover}
        CategoryList={this.props.CategoryList}
        ClientList={this.props.ClientList}
        DepartmentsList={this.props.DepartmentsList}
        ProgramList={this.props.ProgramList}
        ProjectList={this.props.ProjectList}
        ReimbursementRequestList={this.props.ReimbursementRequestList}
        ReimbursementItemsList={this.props.ReimbursementItemsList}
        ReimbursementRequestSettingsList={this.props.ReimbursementRequestSettingsList}
        SubcategoryList={this.props.SubcategoryList}
        TasksList={this.props.TasksList}
      />;
    }

    else {

      const {
        hasTeamsContext,

      } = this.props;
      const deleteIcon: IIconProps = { iconName: 'Delete' };
      const addIcon: IIconProps = { iconName: 'CircleAdditionSolid' };
      const attachmentIcon: IIconProps = { iconName: 'Attach' };



      return (
        <section className={`${styles.reimbursementFormWp} ${hasTeamsContext ? styles.teams : ''}`}>

          <div className={styles.borderBox}>
            <div className={styles.MaterialRequestHeading}>{"REIMBURSEMENT REQUEST FORM"}</div>


            {/* Employee Details */}
            <div className={styles.onediv}>
              <div className={styles.EmployeeDetails}>{"Employee Details"}</div>
              <hr></hr>
              <br></br>

              <div className={styles.employeeDisplay}>
                <div className={styles.fieldwrapper}>
                  <div className={styles.fieldlabel}>Name</div>
                  <div className={styles.colon}>:</div>
                  <div className={styles.fieldoutput}>{this.state.displayName}</div>
                </div>

                <div className={styles.fieldwrapper}>
                  <div className={styles.fieldlabel}>Employee ID</div>
                  <div className={styles.colon}>:</div>
                  <div className={styles.fieldoutput}>{this.state.employeeID}</div>
                </div>
              </div>

              <div className={styles.employeeDisplay}>
                <div className={styles.fieldwrapper}>
                  <div className={styles.fieldlabel}>Department</div>
                  <div className={styles.colon}>:</div>
                  <div className={styles.fieldoutput}>{this.state.department}</div>
                </div>

                <div className={styles.fieldwrapper}>
                  <div className={styles.fieldlabel}>Designation</div>
                  <div className={styles.colon}>:</div>
                  <div className={styles.fieldoutput}>{this.state.designation}</div>
                </div>
              </div>
            </div>


            {/*  Expense Details */}
            <div className={styles.EmployeeDetails}>{" Expense Details"}</div>
            <hr></hr>

            <div className={styles.ProjectWrapper}>
              <Label className={styles.ProjectLabelOne} required={true}> Project</Label>
              <div className={styles.colonOne}>:</div>
              <Dropdown
                className={styles.ProjectInputOne}
                placeholder="Select One"
                options={this.state.project}
                onChange={this.getProjectChange}
                selectedKey={this.state.getProject.key}
              />
            </div>

            <div className={styles.ProjectWrapper}>
              <Label className={styles.ProjectLabelOne} required={true} >Purpose</Label>
              <div className={styles.colonOne}>:</div>
              <TextField
                className={styles.ProjectInputOne}
                value={this.state.purpose}
                onChange={this.onChangePurpose}
              />
            </div>


            <div className={styles.tabledisplay}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th className={styles.tablediv}>#</th>
                    <th className={styles.tablediv}>Date</th>
                    <th className={styles.tablediv}>Category</th>
                    <th className={styles.tablediv}>Sub Category</th>
                    <th className={styles.tablediv}>Expense Details</th>
                    <th className={styles.tablediv}>Amount</th>
                    <th className={styles.iconButton}></th>
                    <th className={styles.iconButton}></th>
                    <th className={styles.iconButton}></th>
                  </tr>
                </thead>

                <tbody>
                  {this.state.rows.map((row: any, index: any) => (
                    <tr key={index}>
                      <td className={styles.tablediv}>{index + 1}</td>

                      {/* <td className={styles.tablediv}> */}
                      <td className={`${styles.tablediv} ${styles.datepickercolumn}`}>
                        <DatePicker
                          placeholder="Select a date..."
                          ariaLabel="Select a date"
                          onSelectDate={(date) => {
                            const updatedRows = [...this.state.rows];
                            updatedRows[index].datepicker = date;
                            this.setState({ rows: updatedRows });
                          }}
                          value={row.datepicker}
                          formatDate={this.onFormatDate}
                          className={styles.dropdownpadding}
                        />
                      </td>
                      <td className={styles.tablediv}>
                        <tr>
                          <Dropdown
                            placeholder="Select Category"
                            required={true}
                            options={this.state.category}
                            selectedKey={row.selectedCategory.key}
                            onChange={(event, option) => this.onCategoryChange(event, option, index)} // Call onCategoryChange
                            // className={styles.dropdownpadding}
                            className={`${styles.dropdownpadding} ${styles.fixedWidthDropdown}`}
                          // style={{ width: '240px' }}
                          />
                        </tr>

                      </td>
                      <td className={styles.tablediv}>
                        <tr>
                          <Dropdown
                            placeholder="Select Subcategory"
                            required={true}
                            options={row.subcategory}
                            selectedKey={row.getSubcategory.key}
                            onChange={(event, option) => {
                              const updatedRows = [...this.state.rows];
                              updatedRows[index].getSubcategory = option || { key: "", text: "" };
                              this.setState({ rows: updatedRows });
                            }}
                            // className={styles.dropdownpadding}
                            className={`${styles.dropdownpadding} ${styles.fixedWidthDropdown}`}
                          />
                        </tr>
                      </td>
                      <td className={styles.tablediv}>
                        <TextField
                          // required={true}
                          placeholder=" "
                          onChange={(event, newValue) => {
                            const updatedRows = [...this.state.rows];
                            updatedRows[index].expenseDetails = newValue || '';
                            this.setState({ rows: updatedRows });
                          }}
                          value={row.expenseDetails}
                          className={styles.dropdownpadding}
                        />

                      </td>
                      <td className={styles.tablediv}>
                        <TextField
                          required={true}
                          placeholder="0"
                          type="text"
                          onChange={(event) => this.handleChangeAmount(event, index)}
                          value={row.amount}
                          style={{ textAlign: 'right' }}
                          className={styles.dropdownpadding}
                        />

                      </td>
                      <td>
                        <IconButton
                          iconProps={addIcon}
                          ariaLabel="Add item"
                          // disabled={!this.state.isQuantityEntered}
                          onClick={this.addRow}
                          className={styles.iconButton}
                        />
                      </td>
                      <td>
                        <IconButton
                          iconProps={deleteIcon}
                          ariaLabel="Delete item"
                          className={styles.iconButton}
                          onClick={() => this.deleteRow(index)}
                          disabled={this.state.rows.length === 1}
                        />
                      </td>
                      <td>
                        {/* Input element for file selection */}
                        <input
                          type="file"
                          multiple={true}
                          style={{ display: 'none' }}
                          onChange={this.onFileChange}
                          ref={(fileInput) => (this.fileInput = fileInput)} // Store reference for programmatic click
                        />
                        {/* Button to trigger file input click */}
                        <IconButton
                          iconProps={attachmentIcon}
                          ariaLabel="Attach files"
                          onClick={() => this.fileInput.click()} // Trigger file input click
                          className={styles.iconButton}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>


            <div className={styles.totaldisplay}>
              <Label className={styles.totaldiv}>Total</Label>
              <div className={styles.amount}>
                {this.state.totalAmount === 0 ? '0' : this.state.totalAmount.toFixed(2)}
              </div>
            </div>




            <div className={styles.ProjectWrapper}>
              <Label className={styles.ProjectLabelOne}>Advance</Label>
              <div className={styles.colonOne}>:</div>
              <TextField
                // required={true}
                placeholder="0"
                type='number'
                onChange={this.onChangeAdvanceAmount}
                style={{ textAlign: 'right' }}
              // className={styles.textfieldAlign}
              />
            </div>

            {/* <div className={styles.ProjectWrapper}>
              <Label className={styles.ProjectLabelOne}> Balance</Label>
              <div className={styles.colonOne}>:</div>
              <TextField
                // required={true}
                placeholder=" "
                type="number"
                value={this.state.balanceAmount.toString()}
                onChange={(event, newValue) => this.onChangeAdvanceAmount(event, newValue || '')}
              />
            </div> */}

            <div className={styles.ProjectWrapper}>
              <Label className={styles.ProjectLabelOne}> Balance:</Label>
              <div className={styles.colonOne}>:</div>
              <Label>{this.state.balanceAmount}</Label>
            </div>


            <div className={styles.ProjectWrapper}>
              <Label className={styles.ProjectLabelOne}> Comment</Label>
              <div className={styles.colonOne}>:</div>
              <TextField
                // label="Comment"
                multiline rows={3}
                onChange={this.onChangeComment}
                value={this.state.comment}
                className={styles.commentArea}
              />
            </div>


            {/* approver details */}
            <div className={styles.EmployeeDetails}>{" Approver Details"}</div>
            <hr></hr>

            <div className={styles.reviewerdisplay}>
              <div className={styles.reviewer}>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Reviewer"
                  // personSelectionLimit={1} // Assuming HOSName is a single user
                  defaultSelectedUsers={[this.state.HOSNameEmail]}
                  groupName={""}
                  showtooltip={true}
                  required={false}
                  disabled={false}
                  ensureUser={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  peoplePickerCntrlclassName={"testClass"}

                />
              </div>

              <div className={styles.reviewer}>
                <PeoplePicker
                  context={this.props.context as any}
                  titleText="Approver"
                  // personSelectionLimit={1} // Assuming HOSName is a single user
                  defaultSelectedUsers={[this.state.adminApproverEmail]}
                  groupName={""}
                  showtooltip={true}
                  required={false}
                  disabled={false}
                  ensureUser={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  peoplePickerCntrlclassName={"testClass"}
                />
              </div>
            </div>

            <div className={styles.reuired} >
              {"* All fields are required"}
            </div>



            <div className={styles.btndiv}>
              <PrimaryButton
                text="Submit"
                onClick={this.onSubmitClick}
                disabled={
                  !this.state.getProject.key ||
                  !this.state.purpose ||
                  !this.areRowsComplete()
                }
              />

              <DefaultButton
                text="Cancel"
                onClick={this.onClickCancel}
              />
            </div>

          </div>

          {/* pop up */}
          <div>
            {this.state.isPopupVisible && (
              <Layer>
                <Popup
                  className={styles.root}
                  role="dialog"
                  aria-modal="true"
                  onDismiss={this.hidePopup}
                >
                  <Overlay
                    onClick={this.hidePopup}
                  />
                  <FocusTrapZone>
                    <div
                      role="document"
                      className={styles.content}
                    >
                      <div>
                        Did you want to apply?
                      </div>

                      <div className={styles.popbtndiv}>
                        <PrimaryButton
                          onClick={this.onPopOk}
                          text="Yes"
                          disabled={this.state.isOkButtonDisabled}

                        />
                        <DefaultButton onClick={this.hidePopup} >No </DefaultButton>
                      </div>

                    </div>
                  </FocusTrapZone>
                </Popup>
              </Layer>
            )}
          </div>

        </section>
      );
    }
  }
}
