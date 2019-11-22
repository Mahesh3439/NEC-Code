import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IListFormService, IcreateSpace } from './ICommonMethods';
import { IProjectSpace } from '../../webparts/projectSummary/components/IProjectSummaryProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IListItem } from '../../webparts/projectSummary/components/IProjectSummaryProps';
import { IErrorLog } from '../../webparts/projectSummary/components/IProjectSummarySubmitProps';
import { sp, Web, WebAddResult } from "@pnp/sp";


export interface IGroups {
  IFAdmin?: number,
  Liaison?: number,
  Agency?: number
}

export class ListFormService implements IListFormService {
  private spHttpClient: SPHttpClient;
  private listItemEntityTypeName: string = undefined;
  constructor(spHttpClient: SPHttpClient) {
    this.spHttpClient = spHttpClient;
  }

  public _vSiteGroups: IGroups = {};


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

  public async _creatProjectSpace(crtSpace: IcreateSpace) {
    try {
      let Api = `${crtSpace.context.pageContext.web.absoluteUrl}/_api/web/GetAvailableWebTemplates(lcid=1033)?$filter=Title eq 'ProjectSpace'`;
      return crtSpace.context.spHttpClient.get(Api, SPHttpClient.configurations.v1)
        .then(resp => { return resp.json(); })
        .then(async (response) => {
          let items = response.value[0];
          let templateName = items.Name;
          const postURL: string = `${crtSpace.context.pageContext.web.absoluteUrl}/_api/web/webinfos/add`;
          const spOpts: ISPHttpClientOptions = {
            body: `{
          "parameters":{
                "@odata.type": "#SP.WebInfoCreationInformation",
                "Title": "${crtSpace.Title}", 
                "Url": "${crtSpace.url}",
                "Description": "${crtSpace.Description}",
                "Language": 1033,
                "WebTemplate": "${templateName}",
                "UseUniquePermissions": true
              }
          }`
          };
          return await crtSpace.context.spHttpClient.post(postURL, SPHttpClient.configurations.v1, spOpts)
            .then(async (response: SPHttpClientResponse) => {
              if (response.ok)
                return response.json();
              else {
                const respText = await response.text();
                throw new Error(respText.toString());
              }

            }).then(async res => {
              await this._getGroups(this._vSiteGroups, crtSpace.context.pageContext.site.absoluteUrl);
              await this._assigneUser(this._vSiteGroups, res.ServerRelativeUrl, crtSpace.investorId);
              await this._assigneLicence(res.ServerRelativeUrl, crtSpace.httpReuest, crtSpace.investorEmail);
              return res;
            });
        });
    }
    catch (error) {
      let errorLog = {
        component: "Project Space creteion",
        page: window.location.href,
        Module: "Project Space",
        exception: error
      }
      await this._logError(crtSpace.context.pageContext.site.absoluteUrl, errorLog);

    }
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

  public async _assigneUser(_vSiteGroups: IGroups, siteURL: string, investor: number) {
    let webURL = `https://ttengage.sharepoint.com${siteURL}`;

    let web = new Web(webURL);
    let IFAdmin = _vSiteGroups.IFAdmin;
    let Liaison = _vSiteGroups.Liaison;
    let Agency = _vSiteGroups.Agency;

    let invRollDef: number = 1073741926;


    //Assigning existing groups to Project Space
    web.roleAssignments.add(IFAdmin, 1073741829);
    web.roleAssignments.add(Liaison, 1073741827);
    //Assigning access to the Investor with contribute rights.
    //UserId and roleDefId
    web.roleAssignments.add(investor, invRollDef);

    web.lists.getByTitle("Documents")
    .rootFolder
    .folders
    .add("Project Summary").then(async function(data) {

      let folderURL = `${siteURL}/Shared%20Documents/Project%20Summary`;
      let summaryFolder = web.getFolderByServerRelativeUrl(folderURL);
      let folder = await summaryFolder.getItem();
      await folder.breakRoleInheritance(false);
      await folder.roleAssignments.add(Agency, 1073741827);
      await folder.roleAssignments.add(Liaison, 1073741829);
      await folder.roleAssignments.add(IFAdmin, 1073741827);
        
    }).catch(function(err) {
        console.log(err);        
    });


    let Approvals = web.lists.getByTitle("Approvals");
    let issues = web.lists.getByTitle("Issues");


    let sitePage: string[] = [`${siteURL}/SitePages/ViewIssue.aspx`, `${siteURL}/SitePages/EditApprovalInfo.aspx`, `${siteURL}/SitePages/Roadmap.aspx`];
    for (let pageURL of sitePage) {

      let getPage = web.getFileByServerRelativeUrl(pageURL);
      let page = await getPage.getItem();
      await page.breakRoleInheritance(false);
      await page.roleAssignments.add(Agency, 1073741827);
      await page.roleAssignments.add(Liaison, 1073741829);
      await page.roleAssignments.add(IFAdmin, 1073741827);
      await page.roleAssignments.add(investor, invRollDef);
    }

    /**
   * Breaking inheritance and assigning access to IF Admin, investor,liaison and Approval Agencys
   */
    await Approvals.breakRoleInheritance(false);
    Approvals.roleAssignments.add(Agency, 1073741827);
    Approvals.roleAssignments.add(Liaison, 1073741829);
    Approvals.roleAssignments.add(IFAdmin, 1073741827);
    Approvals.roleAssignments.add(investor, invRollDef);

    await issues.breakRoleInheritance(false);
    issues.roleAssignments.add(Agency, 1073741827);
    issues.roleAssignments.add(Liaison, 1073741829);
    issues.roleAssignments.add(IFAdmin, 1073741827);
    issues.roleAssignments.add(investor, invRollDef);
  }

  /**
   * 
   * @param siteURL to send site url to capture the error log
   * @param flowURL httpRequest URL to tigger MS Flow for assigning the licience 
   * @param investorEmail Investor email id to assigne c\licience
   */

  public async _assigneLicence(siteURL: string, flowURL: string, investorEmail: string) {
    let postData: any = {
      "UserEmailId": investorEmail
    };
    await fetch(flowURL, {
      method: 'POST',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(postData)
    }).then((response) => {
      //alert("flow is triggered for add licence");

    }).catch((error) => {
      let errorLog: IErrorLog = {
        Module: "User licience",
        component: "Common Method",
        page: "",
        exception: error.toString()
      }
      this._logError(siteURL, errorLog);

    });
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


  public async _getGroups(_vSiteGroups: IGroups, webURL: string) {
    let web = new Web(webURL);
    await web.siteGroups.get().then(function (data) {
      for (let group of data) {
        if (group.Title == "IF Admin") {
          _vSiteGroups.IFAdmin = group.Id;
        }
        else if (group.Title == "Liaison Officer") {
          _vSiteGroups.Liaison = group.Id;
        }
        else if (group.Title == "Approval Agencies") {
          _vSiteGroups.Agency = group.Id;
        }
      }
    });

  }

}