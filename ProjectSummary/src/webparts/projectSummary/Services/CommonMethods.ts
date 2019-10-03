import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IListFormService } from './ICommonMethods';
import { IFieldSchema } from '../components/IProjectSummaryProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IListItem } from '../components/IProjectSummaryProps';


export class ListFormService implements IListFormService {
    private spHttpClient: SPHttpClient;
    private listItemEntityTypeName: string = undefined;
    constructor(spHttpClient: SPHttpClient) {
        this.spHttpClient = spHttpClient;
    }

    /**
     * Gets the schema for all relevant fields for a specified SharePoint list form.
     *
     * @param context The absolute Url to the SharePoint site.
     * @param listTitle The server-relative Url to the SharePoint list.     *
     * @returns Promise object represents the array of field schema for all relevant fields for this list form.
     */
    public getlistFields(context: WebPartContext, listTitle: string): Promise<IFieldSchema[]> {
        return new Promise<IFieldSchema[]>((resolve, reject) => {
            const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/GetByTitle(${listTitle})/fields`;
            this.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
                .then((response: SPHttpClientResponse) => {
                    if (response.ok) {
                        return response.json();
                    }
                })
        });
    }

    /**
     * Gets list items from a specified SharePoint list.
     *
     * @param context The absolute Url to the SharePoint site.
     * @param listTitle The server-relative Url to the SharePoint list.     *
     * @returns Promise object represents the array of Items from requested list.
     */
    public _getListitems(context: WebPartContext, listTitle: string) {
        const restApi = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('` + listTitle + `')/items`;
        return context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
          .then(resp => { return resp.json(); });        
    }

    /**
     * Gets list items from a specified SharePoint list.     *
     * @param context The absolute Url to the SharePoint site.
     * @param listTitle The server-relative Url to the SharePoint list.     *
     * @returns Promise object represents the array of Items from requested list.
     */

    public _getListItemEntityTypeName(context: WebPartContext,lsitTitle: string): Promise<string> {
        return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
          if (this.listItemEntityTypeName) {
            resolve(this.listItemEntityTypeName);
            return;
          }
          const restApi = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${lsitTitle}')?$select=ListItemEntityTypeFullName`;
          this.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
            .then(response => {
              return response.json();
            }, (error: any): void => {
              reject(error);
            })
            .then((response: { ListItemEntityTypeFullName: string }): void => {
              this.listItemEntityTypeName = response.ListItemEntityTypeFullName;
              resolve(this.listItemEntityTypeName);
            });
        });
    }


    /**
     * Gets list items from a specified SharePoint list.     *
     * @param context Context of the Webpart to call SPHttpClient.
     * @param listTitle List title to get data from sharepoint list.     
     * @param ItemID list Item ID to get data.
     * @returns Promise object represents the array of Item fields from a list.
     */

    public _getListItem(context: WebPartContext,listTitle: string, ItemId:number){
      const restApi = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listTitle}')/items(${ItemId})`;
      return context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
        .then(resp => { return resp.json(); });      
     
    }

    public _getListItem_etag(context: WebPartContext,listTitle: string, ItemId:number){
      const restApi = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listTitle}')/items(${ItemId})`;
      return context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        let etag = response.headers.get('ETag');
          return etag; });      
     
    }

    public _getloginusergroups(context: WebPartContext){
      const restApi = `${context.pageContext.web.absoluteUrl}/_api/web/currentuser/?$expand=groups`;
      return context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
        .then(resp => { return resp.json(); });  
    }



}