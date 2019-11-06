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

import { SPPermission } from '@microsoft/sp-page-context';
import { SPComponentLoader } from '@microsoft/sp-loader';

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
    // sp.setup({
    //   spfxContext: this.context
    // });


    SPComponentLoader.loadCss(`${this.context.pageContext.site.absoluteUrl}/SiteAssets/GlobalCSS.css`);
    
    SPComponentLoader.loadScript(`${this.context.pageContext.site.absoluteUrl}/SiteAssets/jquery.js`, {
      globalExportsName: 'jQuery'
  }).catch((error) => {

  }).then((): Promise<{}> => {
      return SPComponentLoader.loadScript(`${this.context.pageContext.site.absoluteUrl}/SiteAssets/jquery.MultiFile.js`, {
          globalExportsName: 'jQuery'
      });
  }).catch((error) => {

  });
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const cssUrl: string = `${this.context.pageContext.site.absoluteUrl}/SiteAssets/CustomCSS.css`;
    const JqueryUrl: string = `${this.context.pageContext.site.absoluteUrl}/SiteAssets/Jquery.js`;
    const JSUrl: string = `${this.context.pageContext.site.absoluteUrl}/SiteAssets/NavContol.js`;
    if (cssUrl) {
      // inject the style sheet
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = cssUrl;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", customStyle);
    }
    if (JqueryUrl) {
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customJquery: HTMLScriptElement = document.createElement("script");
      customJquery.src = JqueryUrl;
      customJquery.type = "text/javascript";
      head.insertAdjacentElement("beforeEnd", customJquery);

    }
    if (JSUrl) {
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customJS: HTMLScriptElement = document.createElement("script");
      customJS.src = JSUrl;
      customJS.type = "text/javascript";
      head.insertAdjacentElement("beforeEnd", customJS);
    }

    this.context.placeholderProvider.changedEvent.add(this, () => {
      this._renderPlaceHolders();
    });

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
      // <input id="btnTranslate" type="button" value="Translate" onClick="Translate()" />
      //<div id="google_translate_element" class=""></div>  
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
         
         </div>
          `;
          // this.LoadSiteBreadcrumb(this);

          //this._UpdateProfile();
         
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



  private _trimSuiteBar(): void {

    sp.web.currentUserHasPermissions(PermissionKind.AddListItems).
      then(
        perms => {
          var suiteBar = document.getElementById("O365_MainLink_Settings");
          if (!suiteBar || !perms) return;
          suiteBar.setAttribute("style", "display: block !important");

          console.log(perms);
        });



    //   pnp.sp.web.usingCaching()
    //     .currentUserHasPermissions(PermissionKind.FullMask)
    //     .then(perms => {
    //       var suiteBar = document.getElementById("O365_MainLink_Settings");
    //       if (!suiteBar || !perms) return; //return if no suite bar OR perms not high enough

    //       suiteBar.setAttribute("style", "display: block !important");
    //     });
  }

  private _UpdateProfile():void{

    //WorkEmail

    //let loginName = `${this.context.pageContext.user.loginName}`;
    let loginName = "i:0#.f|membership|s.agency@ttengage.tt";

      // sp.profiles.myProperties.get().then(function(result) {
      //   var props = result.UserProfileProperties;
      //   var propValue = "";
      //   props.forEach(function(prop) {
      //   propValue += prop.Key + " - " + prop.Value + "<br/>";
      //   });
      //   console.log(propValue);
      //   }).catch(function(err) {
      //   console.log("Error: " + err);
      //   });

      sp.profiles.getUserProfilePropertyFor(loginName,"WorkEmail")
      .then((prop:string)=>{
        console.log(prop);
      })
  }

}
