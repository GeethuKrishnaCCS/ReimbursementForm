import * as React from 'react';
import styles from './ReimbursementRequestHOSApprovalForm.module.scss';
import { IReimbursementRequestHOSApprovalFormProps, IReimbursementRequestHOSApprovalFormState } from '../interfaces/IReimbursementRequestHOSApprovalFormProps';
import { ReimbursementRequestHOSApprovalFormService } from '../services';
import * as moment from 'moment';
import { Label, MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import replaceString from 'replace-string';
import Toast from './Toast';



export default class ReimbursementRequestHOSApprovalForm extends React.Component<IReimbursementRequestHOSApprovalFormProps, IReimbursementRequestHOSApprovalFormState, {}> {
  private _service: any;


  public constructor(props: IReimbursementRequestHOSApprovalFormProps) {
    super(props);
    this._service = new ReimbursementRequestHOSApprovalFormService(this.props.context);

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
    this.sendApprovedEmailNotificationToAdminFromHOS = this.sendApprovedEmailNotificationToAdminFromHOS.bind(this);
    this.sendApprovedEmailNotificationToRequestorFromHOS = this.sendApprovedEmailNotificationToRequestorFromHOS.bind(this);
    this.sendRejectEmailNotificationToRequestorFromHOS = this.sendRejectEmailNotificationToRequestorFromHOS.bind(this);
    this.checkHOS = this.checkHOS.bind(this);
    this.getTaskList = this.getTaskList.bind(this);
  }


  public async componentDidMount() {
    await this.getCurrentUser();
    await this.getReimbursementRequestListData();
    await this.getReimbursementItemsList();
    await this.checkHOS();
    await this.getTaskList();

  }

  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    this.setState({ getcurrentuserId: getcurrentuser.Id });
  }

  public checkHOS() {
    if (this.state.getcurrentuserId !== this.state.HOSApproverId) {
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

  public async OnClickApprove() {
    await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      Status: "HOS Approved",
      // HOSApprovalStatus: "HOS Approved",
      HOSApprovalComments: this.state.comment,
      HOSApprovedDate: moment(new Date()).format("DD-MM-YYYY"),
      // FinanceApprovalStatus: "Admin Pending"
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateEvaluation("ReimbursementRequestList", itemsForUpdate, this.state.getItemId, url);

    await this.deleteTaskListItem();


    const dataItem = {
      ReimbursementRequestIDId: this.state.getItemId,
      AssignedToId: this.state.FinanceApproverId,
    };
    this._service.addListItem(dataItem, this.props.TasksList, url).then(async (task: any) => {

      this.setState({ taskListItemId: task.data.Id });
      const taskURL = url + "/SitePages/" + "ReimbursementRequestAdminApprovalForm" + ".aspx?did=" + this.state.getItemId + "&itemid=" + task.data.Id + "&formType=AdminApproval";
      const taskItemtoupdate = {
        TaskTitleWithLink: {
          Description: "-- Admin Approval",
          Url: taskURL,
        }
      }
      await this._service.updateEvaluation(this.props.TasksList, taskItemtoupdate, task.data.Id, url);

      await this.sendApprovedEmailNotificationToRequestorFromHOS(this.props.context);
      await this.sendApprovedEmailNotificationToAdminFromHOS(this.props.context);
      // alert("approved");
      Toast("success", "Successfully Submitted");
      // setTimeout(() => {
      //   window.location.href = url;
      // }, 3000);
    });
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


  public async sendApprovedEmailNotificationToAdminFromHOS(context: any): Promise<void> {
    const AdminApproverIdUserInfo = await this._service.getUser(this.state.FinanceApproverId);
    const AdminApproverEmail = AdminApproverIdUserInfo.Email;

    // const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const url: string = this.props.context.pageContext.web.absoluteUrl;
    const evaluationURL = url + "/SitePages/" + "ReimbursementRequestAdminApprovalForm" + ".aspx?did=" + this.state.getItemId + "&itemid=" + this.state.taskListItemId + "&formType=AdminApproval";

    const HOSInfo = await this._service.getUser(this.state.HOSApproverId);
    const HOSName = HOSInfo.Title;


    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.ReimbursementRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToAdminFromHOS");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceAdminApprover = replaceString(bodyTemplate, '[AdminApprover]', AdminApproverIdUserInfo.Title)
      const replaceRequestedBy = replaceString(replaceAdminApprover, '[RequestedBy]', this.state.RequestedBy)
      const replaceProject = replaceString(replaceRequestedBy, '[Project]', this.state.project)
      const replaceRequestedDate = replaceString(replaceProject, '[RequestedDate]', this.state.RequestedDate)
      const replaceHOSName = replaceString(replaceRequestedDate, '[HOSName]', HOSName)
      const replacedBodyWithLink = replaceString(replaceHOSName, '[Link]', `<a href="${evaluationURL}" target="_blank">Click here</a>`);

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
                address: AdminApproverEmail,
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

  public async sendApprovedEmailNotificationToRequestorFromHOS(context: any): Promise<void> {
    const requestedBy = await this._service.getUser(this.state.RequestedById);
    const requestedEmail = requestedBy.Email;

    const HOSInfo = await this._service.getUser(this.state.HOSApproverId);
    const HOSName = HOSInfo.Title;
    const date = moment(new Date).format("DD-MM-YYYY");

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.ReimbursementRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToRequestorFromHOS");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestedBy)
      const replaceProject = replaceString(replaceRequestedBy, '[Project]', this.state.project)
      const replaceHOSName = replaceString(replaceProject, '[HOSName]', HOSName)
      const replaceDate = replaceString(replaceHOSName, '[Date]', date)


      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replaceDate
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
    await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      // HOSApprovalStatus: "HOSRejected",
      Status: "HOS Rejected",
      HOSApprovalComments: this.state.comment,
      HOSApprovedDate: moment(new Date()).format("DD-MM-YYYY"),
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateEvaluation(this.props.ReimbursementRequestList, itemsForUpdate, this.state.getItemId, url);

    await this.deleteTaskListItem();

    await this.sendRejectEmailNotificationToRequestorFromHOS(this.props.context);
    // alert("rejected");
    Toast("warning", "Successfully Submitted");
    // setTimeout(() => {
    //   window.location.href = url;
    // }, 3000);
  }

  public async sendRejectEmailNotificationToRequestorFromHOS(context: any): Promise<void> {
    const RequestedByUserInfo = await this._service.getUser(this.state.RequestedById);
    const RequestedByApproverEmail = RequestedByUserInfo.Email;

    const HOSInfo = await this._service.getUser(this.state.HOSApproverId);
    const HOSName = HOSInfo.Title;

    const date = moment(new Date).format("DD-MM-YYYY");

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.ReimbursementRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendRejectEmailNotificationToRequestorFromHOS");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestedBy)
      const replaceProject = replaceString(replaceRequestedBy, '[Project]', this.state.project)
      const replaceHOSName = replaceString(replaceProject, '[HOSName]', HOSName)
      const replaceDate = replaceString(replaceHOSName, '[Date]', date)

      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replaceDate
          },
          toRecipients: [
            {
              emailAddress: {
                address: RequestedByApproverEmail,
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


  public render(): React.ReactElement<IReimbursementRequestHOSApprovalFormProps> {

    const {
      hasTeamsContext,

    } = this.props;

    return (
      <section className={`${styles.reimbursementRequestHOSApprovalForm} ${hasTeamsContext ? styles.teams : ''}`}>

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
                    disabled={this.state.isOkButtonDisabled}
                  />

                  <PrimaryButton
                    text="Reject"
                    onClick={this.OnClickReject}
                    disabled={this.state.isOkButtonDisabled}
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

