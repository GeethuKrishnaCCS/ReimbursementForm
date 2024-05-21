import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/attachments";
import { IItem } from "@pnp/sp/items";

export class ReimbursementFormWpService extends BaseService {
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
    public getClientListItems(listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items();
    }
    public getProgramListItems(listname: string, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Program,ID,Client/Title,Client/ID").expand("Client").filter("Client/ID eq '" + id + "'")();
    }
    public getSubcategoryListItems(listname: string, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Subcategory,ID,Category/Category,Category/ID").expand("Category").filter("Category/ID eq '" + id + "'")();
    }
    public getProjectListItems(listname: string, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Project,ID,Program/Title,Program/ID").expand("Program").filter("Program/ID eq '" + id + "'")();
    }
    public addItemRequestForm(data: any, listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    }   
    public async addAttachments(listname: string, id: number, fileName: string, attachment: any): Promise<any> {
        const item: IItem = await this._spfi.web.lists.getByTitle(listname).items.getById(id)
        return await item.attachmentFiles.add(fileName, attachment)
    }
    public updateRequestForm(listname: string, data: any, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }
    public getListItems(listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items();
    }
  
}