import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IFieldSchema } from '../components/IProjectSummaryProps';
import { IListItem } from '../components/IProjectSummaryProps';

export interface IListFormService {
    getlistFields: (context: WebPartContext, listTitle: string) => Promise<IFieldSchema[]>;
    _getListitems(context: WebPartContext, listTitle: string);
    _getListItemEntityTypeName:(context: WebPartContext,lsitTitle: string)=> Promise<string>;
}