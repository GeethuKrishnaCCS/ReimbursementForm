import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/attachments";

export class ReimbursementRequestHOSApprovalFormService extends BaseService {
    private _spfi: SPFI;
    constructor(context: WebPartContext) {
        super(context);
        this._spfi = getSP(context);
    }


    public async getUser(userId: number): Promise<any> {
        return this._spfi.web.getUserById(userId)();
    }
    public async getCurrentUser(): Promise<any> {
        return this._spfi.web.currentUser();
    }
    public getItemSelectExpandFilter(siteUrl: string, listname: string, select: string, expand: string, filter: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select(select)
            .expand(expand)
            .filter(filter)()
    }
    public getItemSelectExpand(siteUrl: string, listname: string, select: string, expand: string): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items
            .select(select)
            .expand(expand)
            ()
    }
    public updateEvaluation(listname: string, data: any, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
    public deleteItemById(siteUrl: string, listname: string, itemid: number): Promise<any> {
        return this._spfi.web.getList(siteUrl + "/Lists/" + listname).items.getById(itemid).delete();
    }
    public addListItem(data: any, listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    } 
    public getListItems(listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items();
    }
}