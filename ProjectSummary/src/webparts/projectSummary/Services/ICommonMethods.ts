import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IFieldSchema } from '../components/IProjectSummaryProps';
import { IListItem } from '../components/IProjectSummaryProps';
import { Web } from "sp-pnp-js";

export interface IListFormService {
    getlistFields: (context: WebPartContext, listTitle: string) => Promise<IFieldSchema[]>;
    _getListitems(context: WebPartContext, listTitle: string);
    _getListItemEntityTypeName:(context: WebPartContext,lsitTitle: string)=> Promise<string>;
    _getListItem(contet: WebPartContext, listTitle:string,ItemId: number);
    _getloginusergroups(context: WebPartContext);
    _getListItem_etag(contet: WebPartContext, listTitle:string,ItemId: number);
    
}