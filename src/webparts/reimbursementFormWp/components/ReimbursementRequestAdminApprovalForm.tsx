import * as React from 'react';
import styles from './ReimbursementRequestAdminApprovalForm.module.scss';
import { ReimbursementRequestAdminApprovalFormService } from '../services';
import { Label, MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { IReimbursementRequestAdminApprovalFormProps, IReimbursementRequestAdminApprovalFormState } from '../interfaces/IReimbursementRequestAdminApprovalFormProps';
import * as moment from 'moment';
import replaceString from 'replace-string';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import Toast from './Toast';


export default class ReimbursementRequestAdminApprovalForm extends React.Component<IReimbursementRequestAdminApprovalFormProps, IReimbursementRequestAdminApprovalFormState, {}> {
  private _service: any;


  public constructor(props: IReimbursementRequestAdminApprovalFormProps) {
    super(props);
    this._service = new ReimbursementRequestAdminApprovalFormService(this.props.context);

    this.state = {
      getItemId: "",
      getcurrentuserId: null,
      reimbursementRequestData: [],
      RequestedBy: "",
      RequestedDate: "",
      client: "",
      program: "",
      project: "",
      department: "",
      advanceAmount: "",
      balanceAmount: "",
      reimbursementRequestDataId: null,
      reimbursementListData: [],
      totalAmount: "",
      RequestorComments: "",
      comment: "",
      FinanceApproverId: null,
      taskListItemId: null,
      HOSApproverId: null,
      isOkButtonDisabled: false,
      isTaskIdPresent: "",
      noAccessId: "",
      statusMessageTAskIdNull: "",
      RequestedById: null,



    }
    this.getCurrentUser = this.getCurrentUser.bind(this);
    this.getReimbursementRequestListData = this.getReimbursementRequestListData.bind(this);
    this.getReimbursementItemsList = this.getReimbursementItemsList.bind(this);
    this.onChangeComment = this.onChangeComment.bind(this);
    this.OnClickApprove = this.OnClickApprove.bind(this);
    this.OnClickReject = this.OnClickReject.bind(this);
    this.deleteTaskListItem = this.deleteTaskListItem.bind(this);
    this.sendApprovedEmailNotificationToHOSFromAdmin = this.sendApprovedEmailNotificationToHOSFromAdmin.bind(this);
    this.sendApprovedEmailNotificationToRequestorFromAdmin = this.sendApprovedEmailNotificationToRequestorFromAdmin.bind(this);
    this.sendRejectEmailNotificationToHOSFromAdmin = this.sendRejectEmailNotificationToHOSFromAdmin.bind(this);
    this.sendRejectEmailNotificationToRequestorFromAdmin = this.sendRejectEmailNotificationToRequestorFromAdmin.bind(this);
    this.getTaskList = this.getTaskList.bind(this);
    this.checkAdmin = this.checkAdmin.bind(this);
  }


  public async componentDidMount() {
    await this.getCurrentUser();
    await this.getReimbursementRequestListData();
    await this.getReimbursementItemsList();
    await this.checkAdmin();
    await this.getTaskList();



  }

  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    this.setState({ getcurrentuserId: getcurrentuser.Id });
  }

  public async getReimbursementRequestListData() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const itemId = new URLSearchParams(window.location.search).get('did');
    this.setState({ getItemId: itemId });

    const reimbursementRequestData = await this._service.getItemSelectExpandFilter(
      url, this.props.ReimbursementRequestList,
      "*,Project/ID,Project/Project", "Project",
      `Id eq ${itemId}`);
    // console.log('reimbursementRequestData: ', reimbursementRequestData);
    this.setState({ reimbursementRequestData: reimbursementRequestData });
    // console.log('reimbursementRequestData: ', this.state.reimbursementRequestData[0]);
    // console.log(this.state.reimbursementRequestData[0].RequestorComments, "RequestorComments");

    const requestedBy = await this._service.getUser(this.state.reimbursementRequestData[0].AuthorId);
    const RequestedBy = requestedBy.Title;

    const date = this.state.reimbursementRequestData[0].Created
    const dateformatted = moment(date).format("DD-MM-YYYY");

    this.setState({
      reimbursementRequestDataId: this.state.reimbursementRequestData[0].Id,
      RequestedBy: RequestedBy,
      RequestedById: this.state.reimbursementRequestData[0].AuthorId,
      RequestedDate: dateformatted,
      // client: this.state.reimbursementRequestData[0].Client.Client,
      // program: this.state.reimbursementRequestData[0].Program.Program,
      project: this.state.reimbursementRequestData[0].Project.Project,
      department: this.state.reimbursementRequestData[0].Department,
      advanceAmount: this.state.reimbursementRequestData[0].AdvanceAmountNumber,
      balanceAmount: this.state.reimbursementRequestData[0].BalanceAmountNumber,
      totalAmount: this.state.reimbursementRequestData[0].TotalAmountNumber,
      RequestorComments: this.state.reimbursementRequestData[0].RequestorComments,
      FinanceApproverId: this.state.reimbursementRequestData[0].FinanceApproverId,
      HOSApproverId: this.state.reimbursementRequestData[0].HOSApproverId,
    });

  }

  public async getReimbursementItemsList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const reimbursementListData = await this._service.getItemSelectExpandFilter(
      url,
      this.props.ReimbursementItemsList,
      "*, ReimbursementRequestID/ID,ReimbursementRequestID/Title, Category/ID,Category/Category, Subcategory/ID,Subcategory/Subcategory ",
      "ReimbursementRequestID, Category, Subcategory ",
      `ReimbursementRequestID/ID eq ${this.state.reimbursementRequestDataId}`
    );

    this.setState({
      reimbursementListData: reimbursementListData,

    });
    // console.log('reimbursementListData: ', this.state.reimbursementListData);
  }

  public onChangeComment(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, Comment: string) {
    this.setState({ comment: Comment });
  }

  public checkAdmin() {
    if (this.state.getcurrentuserId !== this.state.FinanceApproverId) {
      this.setState({
        noAccessId: "false",
        statusMessageTAskIdNull: "AccessDenied"
      });
    } else {
      this.setState({ noAccessId: "true" });
    }
  }
  public async getTaskList() {
    const taskItemid = new URLSearchParams(window.location.search).get('itemid');

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const taskListData: any[] = await this._service.getItemSelectExpandFilter(
      url,
      this.props.TasksList,
      "ID, TaskTitleWithLink, ReimbursementRequestID/ID",
      "ReimbursementRequestID",
      `ID eq ${taskItemid}`
    );

    if (taskListData.length === 0) {
      this.setState({
        isTaskIdPresent: "false",
        statusMessageTAskIdNull: "Alreadycheckedtherequest"
      });
    } else {
      this.setState({ isTaskIdPresent: "true" });
    }
  }

  public async OnClickApprove() {
    // await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      Status: "Finance Approved",
      // FinanceApprovalStatus: "Admin Approved",
      FinanceApproverComments: this.state.comment,
      FinanceApproverDate: moment(new Date()).format("DD-MM-YYYY"),
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateEvaluation(this.props.ReimbursementRequestList, itemsForUpdate, this.state.getItemId, url);

    await this.deleteTaskListItem();
    await this.sendApprovedEmailNotificationToHOSFromAdmin(this.props.context);
    await this.sendApprovedEmailNotificationToRequestorFromAdmin(this.props.context);
    Toast("success", "Successfully approved!");
    setTimeout(() => {
      window.location.href = url;
    }, 3000);
  }

  public async deleteTaskListItem() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const taskListData = await this._service.getItemSelectExpandFilter(
      url,
      this.props.TasksList,
      "ID, ReimbursementRequestID/ID",
      "ReimbursementRequestID",
      `ReimbursementRequestID/ID eq ${this.state.getItemId}`
    );

    const taskIdItem = taskListData[0].ID;
    await this._service.deleteItemById(url, this.props.TasksList, taskIdItem);
  }

  public async sendApprovedEmailNotificationToHOSFromAdmin(context: any): Promise<void> {
    const HOSApproverIdUserInfo = await this._service.getUser(this.state.HOSApproverId);
    const HOSApproverEmail = HOSApproverIdUserInfo.Email;
    const date = moment(new Date).format("DD-MM-YYYY");

    const FinanceApproverIdUserInfo = await this._service.getUser(this.state.FinanceApproverId);

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.ReimbursementRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToHOSFromAdmin");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceHOSApprover = replaceString(bodyTemplate, '[HOSApprover]', HOSApproverIdUserInfo.Title)
      const replaceProject = replaceString(replaceHOSApprover, '[Project]', this.state.project)
      const replaceApprovedBy = replaceString(replaceProject, '[ApprovedBy]', FinanceApproverIdUserInfo.Title)
      const replacedate = replaceString(replaceApprovedBy, '[Date]', date)


      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replacedate
          },
          toRecipients: [
            {
              emailAddress: {
                address: HOSApproverEmail,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }

  public async sendApprovedEmailNotificationToRequestorFromAdmin(context: any): Promise<void> {
    const requestedBy = await this._service.getUser(this.state.RequestedById);
    const requestedEmail = requestedBy.Email;

    const date = moment(new Date).format("DD-MM-YYYY");

    const FinanceApproverIdUserInfo = await this._service.getUser(this.state.FinanceApproverId);

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.ReimbursementRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToRequestorFromAdmin");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestedBy)
      const replaceProject = replaceString(replaceRequestedBy, '[Project]', this.state.project)
      const replaceApprovedBy = replaceString(replaceProject, '[ApprovedBy]', FinanceApproverIdUserInfo.Title)
      const replacedate = replaceString(replaceApprovedBy, '[Date]', date)

      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replacedate
          },
          toRecipients: [
            {
              emailAddress: {
                address: requestedEmail,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }

  public async OnClickReject() {
    // await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      Status: "Finance Rejected",
      // FinanceApprovalStatus: "Admin Rejected",
      FinanceApproverComments: this.state.comment,
      FinanceApproverDate: moment(new Date()).format("DD-MM-YYYY"),
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateEvaluation(this.props.ReimbursementRequestList, itemsForUpdate, this.state.getItemId, url);

    await this.deleteTaskListItem();
    
    await this.sendRejectEmailNotificationToHOSFromAdmin(this.props.context);
    await this.sendRejectEmailNotificationToRequestorFromAdmin(this.props.context);
    Toast("warning", "Rejected");
    setTimeout(() => {
      window.location.href = url;
    }, 3000);
  }

  public async sendRejectEmailNotificationToHOSFromAdmin(context: any): Promise<void> {
    const HOSApproverIdUserInfo = await this._service.getUser(this.state.HOSApproverId);
    const HOSApproverEmail = HOSApproverIdUserInfo.Email;
    const date = moment(new Date).format("DD-MM-YYYY");
    const FinanceApproverIdUserInfo = await this._service.getUser(this.state.FinanceApproverId);

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.ReimbursementRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendRejectEmailNotificationToHOSFromAdmin");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceHOSApprover = replaceString(bodyTemplate, '[HOSApprover]', HOSApproverIdUserInfo.Title)
      const replaceProject = replaceString(replaceHOSApprover, '[Project]', this.state.project)
      const replaceApprovedBy = replaceString(replaceProject, '[ApprovedBy]', FinanceApproverIdUserInfo.Title)
      const replacedate = replaceString(replaceApprovedBy, '[Date]', date)


      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replacedate
          },
          toRecipients: [
            {
              emailAddress: {
                address: HOSApproverEmail,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }

  public async sendRejectEmailNotificationToRequestorFromAdmin(context: any): Promise<void> {
    const requestedBy = await this._service.getUser(this.state.RequestedById);
    const requestedEmail = requestedBy.Email;
    const date = moment(new Date).format("DD-MM-YYYY");

    const FinanceApproverIdUserInfo = await this._service.getUser(this.state.FinanceApproverId);

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.ReimbursementRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendRejectEmailNotificationToRequestorFromAdmin");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestedBy)
      const replaceProject = replaceString(replaceRequestedBy, '[Project]', this.state.project)
      const replaceApprovedBy = replaceString(replaceProject, '[ApprovedBy]', FinanceApproverIdUserInfo.Title)
      const replacedate = replaceString(replaceApprovedBy, '[Date]', date)

      // const HOSInfo = await this._service.getUser(this.state.materialRequestData.HOSApproverId);
      // const HOSName = HOSInfo.Title;

      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replacedate
          },
          toRecipients: [
            {
              emailAddress: {
                address: requestedEmail,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }


  public render(): React.ReactElement<IReimbursementRequestAdminApprovalFormProps> {

    const {
      hasTeamsContext,

    } = this.props;

    return (
      <section className={`${styles.reimbursementRequestAdminApprovalForm} ${hasTeamsContext ? styles.teams : ''}`}>

        <div className={styles.borderBox}>
          <div>
            {this.state.noAccessId === "false" &&
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {this.state.statusMessageTAskIdNull}
              </MessageBar>
            }
          </div>

          <div>
            {this.state.isTaskIdPresent === "false" && this.state.noAccessId === "true" &&
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {this.state.statusMessageTAskIdNull}
              </MessageBar>
            }
          </div>

          <div>
            {this.state.isTaskIdPresent === "true" && this.state.noAccessId === "true" &&
              <div>
                <div className={styles.MaterialRequestHeading}>{"Reimbursement Request"}</div>

                <div className={styles.onediv}>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Requested By</div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.RequestedBy}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Requested Date </div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.RequestedDate}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Project </div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.project}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Department</div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.department}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Advance Amount</div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.advanceAmount}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Balance Amount </div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.balanceAmount}</div>
                  </div>

                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Requestor Comments </div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.RequestorComments}</div>
                  </div>

                </div>

                <div>
                  <table className={`${styles.table} ${styles.tablethtddiv}`}>
                    <thead>
                      <tr>
                        <th className={styles.tablethtddiv}>SL No</th>
                        <th className={styles.tablethtddiv}>Expense Date</th>
                        <th className={styles.tablethtddiv}>Category</th>
                        <th className={styles.tablethtddiv}>Subcategory</th>
                        <th className={styles.tablethtddiv}>Expense</th>
                        <th className={styles.tablethtddiv}>Amount</th>

                      </tr>
                    </thead>
                    <tbody>
                      {this.state.reimbursementListData.map((reimbursementItem: any, index: any) => (
                        <tr key={index}>
                          <td className={styles.tablethtddiv}>{index + 1}</td>
                          <td className={styles.tablethtddiv}>{reimbursementItem.ExpenseDate}</td>
                          <td className={styles.tablethtddiv}>{reimbursementItem.Category.Category}</td>
                          <td className={styles.tablethtddiv}>{reimbursementItem.Subcategory.Subcategory}</td>
                          <td className={styles.tablethtddiv}>{reimbursementItem.ExpenseDetails}</td>
                          <td className={styles.tablethtddiv}>{reimbursementItem.Amount}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>


                <div className={styles.totaldisplay}>
                  <Label className={styles.totaldiv}>Total</Label>
                  <div className={styles.amount}>  {this.state.totalAmount} </div>
                </div>

                <div>
                  <TextField
                    label="Comment"
                    multiline rows={3}
                    onChange={this.onChangeComment}
                    value={this.state.comment}
                  // className={styles.commentArea}
                  />
                </div>

                <div className={styles.btndiv}>
                  <PrimaryButton
                    text="Approve"
                    onClick={this.OnClickApprove}
                  // disabled={this.state.isOkButtonDisabled}
                  />

                  <PrimaryButton
                    text="Reject"
                    onClick={this.OnClickReject}
                  // disabled={this.state.isOkButtonDisabled}
                  />
                </div>
              </div>
            }
          </div>
        </div>



      </section>
    );
  }
}

