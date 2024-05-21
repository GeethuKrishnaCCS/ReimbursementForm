import * as React from 'react';
import { IReimbursementRequestViewTableProps, IReimbursementRequestViewTableState } from '../interfaces/IReimbursementRequestViewTableProps';
import styles from './ReimbursementRequestViewTable.module.scss';
import { ReimbursementRequestViewTableService } from '../services';
// import * as moment from 'moment';
import { Pagination } from '@pnp/spfx-controls-react/lib/pagination';
import * as _ from 'lodash';
import { DetailsList, DetailsListLayoutMode, IColumn, MessageBar, MessageBarType, SelectionMode } from '@fluentui/react';

export default class ReimbursementRequestViewTable extends React.Component<IReimbursementRequestViewTableProps, IReimbursementRequestViewTableState, {}> {
  private _service: any;


  public constructor(props: IReimbursementRequestViewTableProps) {
    super(props);
    this._service = new ReimbursementRequestViewTableService(this.props.context);

    this.state = {
      noItems: "",
      statusMessageNoItems: "",
      materialDatas: [],
      groups: [],
      expandedItems: [],
      statusGroups: [],
      combinedGroups: [],
      adminApproverId: null,
      departmentName: '',
      HOSName: null,
      Departmentslist: [],
      getcurrentuserId: null,
      currentPage: 1,
      itemsPerPage: 5,
    }

    this.getCurrentUser = this.getCurrentUser.bind(this);
    this.getAdminApprover = this.getAdminApprover.bind(this);
    this.getDepartmentsList = this.getDepartmentsList.bind(this);
    this.handlePageChange = this.handlePageChange.bind(this);
    this.getDocumentIndexItems = this.getDocumentIndexItems.bind(this);

  }


  public async componentDidMount() {
    await this.getAdminApprover();
    await this.getCurrentUser();
    await this.getDepartmentsList();
    await this.getDocumentIndexItems();


  }

  public async getDocumentIndexItems() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const currentUserId = this.state.getcurrentuserId;
    const isAdmin = currentUserId === this.state.adminApproverId;

    const isHOS = this.state.Departmentslist.some((dept: any) => dept.HOSNameId === currentUserId);

    let reimbursementRequestData = await this._service.getItemSelectExpand(url, this.props.ReimbursementRequestList,
      "*,Project/ID,Project/Project,Client/ID,Client/Client,Program/ID,Program/Program", "Project,Client,Program");

    const reimbursementItemListVal = await this._service.getItemSelectExpand(
      url,
      this.props.ReimbursementItemsList,
      "*, ReimbursementRequestID/ID,ReimbursementRequestID/Title,Category/ID,Category/Category,Subcategory/ID,Subcategory/Subcategory",
      "ReimbursementRequestID,Category, Subcategory"
    );

    if (isAdmin) {
    } else if (isHOS) {
      reimbursementRequestData = reimbursementRequestData.filter((data: any) => data.AuthorId === currentUserId ||
        reimbursementRequestData.filter((item: any) => item.HOSApproverId === currentUserId));
    } else {
      reimbursementRequestData = reimbursementRequestData.filter((data: any) => data.AuthorId === currentUserId);
    }

    const expandedItems: any[] = [];

    for (const data of reimbursementRequestData) {
      const reimbursementItemList = reimbursementItemListVal.filter((d: any) => d.ReimbursementRequestID.ID === data.Id);

      for (const reimbursementItem of reimbursementItemList) {
        expandedItems.push({
          reimbursementRequestDataID: data.Id,
          reimbursementRequestCode: data.ReimbursementRequestCode,
          requestedDate: data.RequestedDate,
          expenseDate: reimbursementItem.ExpenseDate,
          category: reimbursementItem.Category.Category,
          subcategory: reimbursementItem.Subcategory.Subcategory,
          status: data.Status,
          amount: reimbursementItem.Amount,
          expenseDetails: reimbursementItem.ExpenseDetails,
          totalAmountNumber: data.TotalAmountNumber,
        });
      }
    }


    const groupedByReimbursementRequest = _.groupBy(expandedItems, 'reimbursementRequestCode');

    let cumulativeCount = 0;
    const groups = await Promise.all(_.map(groupedByReimbursementRequest, async (reimbursementItems: any, reimbursementRequestCode: any) => {
      const groupedByStatus = _.groupBy(reimbursementItems, 'status');
      const statusGroups = _.map(groupedByStatus, (statusItems: any, status: any) => {
        const statusCount = statusItems.length;
        const startIndex = cumulativeCount;
        cumulativeCount += statusCount;
        return {
          key: `${reimbursementRequestCode}_${status}`,
          name: status,
          startIndex: startIndex,
          count: statusCount,
          level: 1,
        };
      });

      return {
        key: reimbursementRequestCode,
        name: reimbursementRequestCode,
        startIndex: cumulativeCount - reimbursementItems.length,
        count: reimbursementItems.length,
        children: statusGroups,
      };
    }));

    this.setState({
      expandedItems: expandedItems,
      groups: groups,

    });
    if (this.state.expandedItems.length === 0) {
      this.setState({
        noItems: "false",
        statusMessageNoItems: 'No items to display'
      });
    } else {
      this.setState({ noItems: "true" });
    }
  }

  public async getAdminApprover() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const AdminApproverlistItem = await this._service.getListItems(this.props.AdminApprover, url)
    const ApproverIdUserInfo = await this._service.getUser(AdminApproverlistItem[0].AdminApproverId);
    this.setState({ adminApproverId: ApproverIdUserInfo.Id });
  }

  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    this.setState({ getcurrentuserId: getcurrentuser.Id });
  }

  public async getDepartmentsList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const DepartmentslistItem = await this._service.getListItems(this.props.DepartmentsList, url)
    this.setState({ Departmentslist: DepartmentslistItem });

    DepartmentslistItem.map((Item: any) => {
      const departmentName = Item.Title;
      const HOSName = Item.HOSNameId;

      this.setState({
        departmentName: departmentName,
        HOSName: HOSName,
      });
    })
  }

  public handlePageChange = (pageNumber: number) => {
    this.setState({
      currentPage: pageNumber
    });
  };

  // private _onRenderGroupFooter: IDetailsGroupRenderProps['onRenderFooter'] = props => {
  //   if (props) {
  //     console.log(props.group.children);
  //     console.log(props);
  //     console.log(props.groups[0].data);
  //     return (
  //       <div >
  //         <em>{`Total ${props.group["totalsum"]}`}</em>
  //       </div>
  //     );
  //   }

  //   return null;
  // };

  public render(): React.ReactElement<IReimbursementRequestViewTableState> {

    const {
      hasTeamsContext,

    } = this.props;

    const columns: IColumn[] = [
      // {
      //   key: 'column0',
      //   name: 'reimbursementRequest Code ',
      //   fieldName: 'reimbursementRequestCode',
      //   minWidth: 210,
      //   maxWidth: 350,
      //   isRowHeader: true,
      //   isResizable: true,
      //   isSorted: true,
      //   isSortedDescending: false,
      //   sortAscendingAriaLabel: 'Sorted A to Z',
      //   sortDescendingAriaLabel: 'Sorted Z to A',
      //   data: 'string',
      //   isPadded: true,
      // },
      {
        key: 'column1',
        name: 'Requested Date',
        fieldName: 'requestedDate',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column2',
        name: 'Expense Date',
        fieldName: 'expenseDate',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column3',
        name: 'Category',
        fieldName: 'category',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column4',
        name: 'Subcategory',
        fieldName: 'subcategory',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'string',

        isPadded: true,
      },
      {
        key: 'column5',
        name: 'Expense Details',
        fieldName: 'expenseDetails',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column6',
        name: 'Amount',
        fieldName: 'amount',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'number',
      },


    ];

    const indexOfLastItem = this.state.currentPage * this.state.itemsPerPage;
    const indexOfFirstItem = indexOfLastItem - this.state.itemsPerPage;
    const currentGroups = this.state.groups.slice(indexOfFirstItem, indexOfLastItem);

    const totalPages = Math.ceil(this.state.groups.length / this.state.itemsPerPage);


    return (
      <section className={`${styles.reimbursementRequestViewTable} ${hasTeamsContext ? styles.teams : ''}`}>

        <div className={styles.borderBox}>

          <div className={styles.MaterialRequestHeading}>{"Reimbursement Request"}</div>

          <div>
            {this.state.noItems === "false" &&
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {this.state.statusMessageNoItems}
              </MessageBar>
            }
          </div>
          <div>
            {this.state.noItems === "true" &&
              <>
                <DetailsList
                  items={this.state.expandedItems}
                  columns={columns}
                  setKey='set'
                  layoutMode={DetailsListLayoutMode.justified}
                  isHeaderVisible={true}
                  selectionMode={SelectionMode.none}
                  groups={currentGroups}
                // groups={this.state.groups}
                //  groupProps={{
                //  onRenderFooter: this._onRenderGroupFooter,
                // }}
                />

                <Pagination
                  totalPages={totalPages}
                  currentPage={this.state.currentPage}
                  onChange={this.handlePageChange}
                />
              </>
            }

          </div>
        </div>


      </section>
    );
  }
}

