import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'NecExtensionApplicationCustomizerStrings';
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

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const cssUrl: string = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/CustomCSS.css`;
    const JqueryUrl: string = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/Jquery.js`;
    const JSUrl: string = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/NavContol.js`;
    if (cssUrl) {
      // inject the style sheet
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = cssUrl;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", customStyle);
    }
    if(JqueryUrl)
    {
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customJquery: HTMLScriptElement = document.createElement("script");
      customJquery.src = JqueryUrl;      
      customJquery.type = "text/javascript";
      head.insertAdjacentElement("beforeEnd", customJquery);

    } 
    if(JSUrl)
    {
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customJS: HTMLScriptElement = document.createElement("script");
      customJS.src = JSUrl;      
      customJS.type = "text/javascript";
      head.insertAdjacentElement("beforeEnd", customJS);

    }  

    return Promise.resolve();
  }
}
