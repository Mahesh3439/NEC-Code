import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'NecExtensionApplicationCustomizerStrings';

//import '../../Commonfiles/Services/customStyles.css';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import pnp, { PermissionKind } from "sp-pnp-js";
import { sp } from "@pnp/sp";


import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient } from '@microsoft/sp-http';
import { IconButton } from 'office-ui-fabric-react';
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');


const LOG_SOURCE: string = 'NecExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INecExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class NecExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<INecExtensionApplicationCustomizerProperties> {

  private _headerPlaceholder: PlaceholderContent;
  private _footerPlaceholder: PlaceholderContent;

  @override
  public onInit(): Promise<void> {
    SPComponentLoader.loadCss(`${this.context.pageContext.site.absoluteUrl}/SiteAssets/GlobalCSS.css`);

    SPComponentLoader.loadScript(`${this.context.pageContext.site.absoluteUrl}/SiteAssets/jquery.js`, {
      globalExportsName: 'jQuery'
    }).catch((error) => {

    }).then(() => {
      SPComponentLoader.loadScript(`${this.context.pageContext.site.absoluteUrl}/SiteAssets/jquery.MultiFile.js`, {
        globalExportsName: 'jQuery'
      });
      // SPComponentLoader.loadScript(`${this.context.pageContext.site.absoluteUrl}/SiteAssets/NavContol.js`, {
      //   globalExportsName: 'getUserGroups'
      // });

      if (!sessionStorage.getItem("loginuser")) {
        $("#O365_MainLink_Settings").parent().hide();
        $(".SiteContent").hide()
      }



    }).catch((error) => {

    });
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    this.context.placeholderProvider.changedEvent.add(this, () => {
      this._renderPlaceHolders();
      this._getUserGroups();
      if (!sessionStorage.getItem("loginuser")) {
        this._validateUser();
      }

    });

    // this.context.application.navigatedEvent.add(this,async ()=>{
    //   await this._getUserGroups();

    // });

    return Promise.resolve();
  }

  private _onDispose(): void {
    console.log('[Breadcrumb._onDispose] Disposed breadcrumb.');
  }

  private _renderPlaceHolders(): void {
    if (!this._headerPlaceholder) {
      this._headerPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._headerPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this.properties) {
        if (this._headerPlaceholder.domElement) {
          this._headerPlaceholder.domElement.innerHTML = `          
         <div class="navbar"> 
         <div class="navbar-header pull-left">
					<a href="${this.context.pageContext.site.absoluteUrl}" class="navbar-brand">
						<small>
							<div class="login-logo">
                <img src="${this.context.pageContext.site.absoluteUrl}/SiteAssets/SiteImages/NECLogo1.png" alt="" ></div>
						</small>
					</a>
        </div>
         <div class="navbar-buttons navbar-header pull-right " role="navigation">
					<ul class="nav ace-nav">
              <li style="border-left:0px;display:block">
                  <div class="login-logo">
                    <img src="${this.context.pageContext.site.absoluteUrl}/SiteAssets/SiteImages/NECLogo.png" alt="" >
                  </div>
              </li>               
             
					</ul>
				</div>  
         
        <a class="SiteContent" title="SiteContent" href="${this.context.pageContext.site.absoluteUrl}/sitepages/sitecontent.aspx" class="ms-Button ms-Button--primary " data-is-focusable="true">
        <div class="ms-Button-flexContainer"><div class="ms-Button-textContainer"><div class="ms-Button-label label-236">Site Content</div></div></div>
        </a>

         </div>
          `;

        }
      }
    }

    if (!this._footerPlaceholder) {
      this._footerPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._footerPlaceholder) {
        console.error('The expected placeholder (Bottom) was not found.');
        return;
      }

      if (this.properties) {
        if (this._footerPlaceholder.domElement) {
          this._footerPlaceholder.domElement.innerHTML = `          
          <div class="footer">
              <div class="footer-inner">
                <div class="footer-content">
                    <span class="bigger-120">                        
                      Copyright 2019 by National Energy     
                    </span>
                    &nbsp; &nbsp;
                </div>
              </div>
          </div>
          `;
          // this.LoadSiteBreadcrumb(this);
        }
      }

    }
  }

  public async _getUserGroups() {

    if (!sessionStorage.getItem("UserGroups")) {
      const restApi = `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser/?$expand=groups`;
      await this.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
        .then(resp => { return resp.json(); })
        .then((data) => {
          sessionStorage.setItem("UserGroups", JSON.stringify(data.Groups));
          this._ControlNavigation();
        });
    }
    else {
      this._ControlNavigation();
    }

  }

  public _ControlNavigation() {
    let userGroups = JSON.parse(sessionStorage.getItem("UserGroups"));

    for (const uGroup of userGroups) {
      if (uGroup.Title == "IF Admin") {
        //document.getElementById('O365_MainLink_Settings').style.display = 'inline-block !important';
        document.getElementsByName("Submit a Project")[0].style.display = "block";
        sessionStorage.setItem("loginuser", "IF Admin");

        return false;
      }
      else if (uGroup.Title == "Investors") {
        document.getElementsByName("Submit a Project")[0].style.display = "block";
        document.getElementById('O365_MainLink_Settings').parentNode[0].style.display = 'none';
        return false;
      }
      else if (uGroup.Title == "Approval Agencies") {
        document.getElementsByName("Agency Dashboard")[0].style.display = "block";
        document.getElementById('O365_MainLink_Settings').parentNode[0].style.display = 'none';
        return false;

      }
    }
  }

  private _validateUser() {
    let pageContex = this.context.pageContext.legacyPageContext;
    let listTitle = pageContex.listTitle;
    let pageURL = pageContex.serverRequestPath;

    if (pageURL.indexOf('/Lists') > 0 || pageURL.indexOf('/Forms') > 0) {
      window.location.href = this.context.pageContext.web.absoluteUrl;
    }



  }

}
