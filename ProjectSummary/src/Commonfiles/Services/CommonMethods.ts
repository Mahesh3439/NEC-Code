import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IListFormService } from './ICommonMethods';
import { IProjectSpace } from '../../webparts/projectSummary/components/IProjectSummaryProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IListItem } from '../../webparts/projectSummary/components/IProjectSummaryProps';
import { IErrorLog } from '../../webparts/projectSummary/components/IProjectSummarySubmitProps';
import { sp, Web, WebAddResult } from "@pnp/sp";



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
  // public getlistFields(context: WebPartContext, listTitle: string): Promise<IFieldSchema[]> {
  //     return new Promise<IFieldSchema[]>((resolve, reject) => {
  //         const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/GetByTitle(${listTitle})/fields`;
  //         this.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
  //             .then((response: SPHttpClientResponse) => {
  //                 if (response.ok) {
  //                     return response.json();
  //                 }
  //             })
  //     });
  // }

  /**
   * Gets list items from a specified SharePoint list.
   *
   * @param context The absolute Url to the SharePoint site.
   * @param listTitle The server-relative Url to the SharePoint list.     *
   * @returns Promise object represents the array of Items from requested list.
   */
  public _getListitems(context: WebPartContext, restApi: string) {
    //const restApi = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('` + listTitle + `')/items`;
    return context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); });
  }

  /**
   * Gets list items from a specified SharePoint list.     *
   * @param context The absolute Url to the SharePoint site.
   * @param listTitle The server-relative Url to the SharePoint list.     *
   * @returns Promise object represents the array of Items from requested list.
   */

  public _getListItemEntityTypeName(context: WebPartContext, lsitTitle: string): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      // if (this.listItemEntityTypeName) {
      //   resolve(this.listItemEntityTypeName);
      //   return;
      // }
      const restApi = `${context.pageContext.site.absoluteUrl}/_api/web/lists/GetByTitle('${lsitTitle}')?$select=ListItemEntityTypeFullName`;
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

  public _getListItem(context: WebPartContext, restApi: string) {
    //const restApi = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${listTitle}')/items(${ItemId})?$select=*,LiaisonOfficer/Id,LiaisonOfficer/EMail&$expand=LiaisonOfficer`;
    return context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); });

  }

  public async _getListItem_etag(context: WebPartContext, listTitle: string, ItemId: number) {
    const restApi = `${context.pageContext.site.absoluteUrl}/_api/web/lists/GetByTitle('${listTitle}')/items(${ItemId})`;
    return context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        let etag = response.headers.get('ETag');
        return etag;
      });

  }

  public _getloginusergroups(context: WebPartContext) {
    const restApi = `${context.pageContext.web.absoluteUrl}/_api/web/currentuser/?$expand=groups`;
    return context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); });
  }

  /**
   * @context: webpart context
   * @siteTitle: ProjectSpace title to create a subsite
   * @siteURL: site url to create a project space.
   * @webtemplate: {8CF9E84A-CD4E-4C2C-847E-2EB55655D939}#ProjectSpace
   */

  public async _creatProjectSpace(context: WebPartContext, siteTitle: string, siteURL: string, investor: number) {
    let Api = `${context.pageContext.web.absoluteUrl}/_api/web/GetAvailableWebTemplates(lcid=1033)?$filter=Title eq 'ProjectSpace'`;
    return context.spHttpClient.get(Api, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(async (response) => {
        let items = response.value[0];
        let templateName = items.Name;
        const postURL: string = `${context.pageContext.web.absoluteUrl}/_api/web/webinfos/add`;
        const spOpts: ISPHttpClientOptions = {
          body: `{
          "parameters":{
                "@odata.type": "#SP.WebInfoCreationInformation",
                "Title": "${siteTitle}", 
                "Url": "${siteURL}",
                "Description": "Projectspace",
                "Language": 1033,
                "WebTemplate": "${templateName}",
                "UseUniquePermissions": true
              }
          }`
        };
        return await context.spHttpClient.post(postURL, SPHttpClient.configurations.v1, spOpts)
          .then((response: SPHttpClientResponse) => {
            return response.json();
          }).then(async res => {
            await this._assigneUser(res.ServerRelativeUrl, investor);
            return res;
          });
      });
  }


  /**
       * 26 - IF Admin Group
       * 25 - Liaison Group
       * 24 - Approval Agencies
       * 69 - Project Investor
       * roleDefId -- 1073741829 -- FullControl
       * roleDefId -- 1073741830 -- Edit
       * roleDefId -- 1073741827 -- Contribute
       */
  public async _assigneUser(siteURL: string, investor: number) {
    let webURL = `https://ttengage.sharepoint.com${siteURL}`;

    let web = new Web(webURL);

    //Assigning existing groups to Project Space
    web.roleAssignments.add(26, 1073741829);
    web.roleAssignments.add(25, 1073741829);
    //Assigning access to the Investor with contribute rights.
    //UserId and roleDefId
    web.roleAssignments.add(investor, 1073741827);


    let Approvals = web.lists.getByTitle("Approvals");
    let issues = web.lists.getByTitle("Issues");


    let sitePage: string[] = [`${siteURL}/SitePages/ViewIssue.aspx`, `${siteURL}/SitePages/EditApprovalInfo.aspx`, `${siteURL}/SitePages/Roadmap.aspx`];
    for (let pageURL of sitePage) {

      let getPage = web.getFileByServerRelativeUrl(pageURL);
      let page = await getPage.getItem();
      await page.breakRoleInheritance(false);
      await page.roleAssignments.add(24, 1073741827);
      await page.roleAssignments.add(25, 1073741829);
      await page.roleAssignments.add(26, 1073741829);
      await page.roleAssignments.add(investor, 1073741827);

    }

    /**
   * Breaking inheritance and assigning access to IF Admin, investor,liaison and Approval Agencys
   */
    await Approvals.breakRoleInheritance(false);
    Approvals.roleAssignments.add(24, 1073741827);
    Approvals.roleAssignments.add(25, 1073741829);
    Approvals.roleAssignments.add(26, 1073741829);
    Approvals.roleAssignments.add(investor, 1073741827);

    await issues.breakRoleInheritance(false);
    issues.roleAssignments.add(24, 1073741827);
    issues.roleAssignments.add(25, 1073741829);
    issues.roleAssignments.add(26, 1073741829);
    issues.roleAssignments.add(investor, 1073741827);


  }

  public async _logError(siteURL: string, errorLog: IErrorLog) {

    let web = new Web(siteURL);
    await web.lists.getByTitle("Exception Log").items.add({
      Title: errorLog.component,
      Page: errorLog.page,
      Module: errorLog.Module,
      Exception: errorLog.exception.toString()
    });
  }
}