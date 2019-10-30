import { WebPartContext } from "@microsoft/sp-webpart-base";


export interface IListFormService {
    //getlistFields: (context: WebPartContext, listTitle: string) => Promise<IFieldSchema[]>;
    _getListitems(context: WebPartContext, APIUrl: string);
    _getListItemEntityTypeName:(context: WebPartContext,lsitTitle: string)=> Promise<string>;
    _getListItem(contet: WebPartContext, apiURL:string);
    _getloginusergroups(context: WebPartContext);
    _getListItem_etag(contet: WebPartContext, listTitle:string,ItemId: number);
    _creatProjectSpace(conte:WebPartContext,siteTitle:string,siteURL:string);

}