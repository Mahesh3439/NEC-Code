import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';

import * as strings from 'PromotionResponseWebPartStrings';
import PromotionResponse from './components/PromotionResponse';
import PromotionResponseNew from './components/PromotionResponseNew';
import PromotionResponseEdit from './components/PromotionResponseEdit';
import { IPromotionResponseProps } from './components/IPromotionResponseProps';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IPromotionResponseWebPartProps {
  description: string;
  FormType: string;
  context:WebPartContext;
  httpRequest:string;
}

export default class PromotionResponseWebPart extends BaseClientSideWebPart<IPromotionResponseWebPartProps> {

  public render(): void {
    let cssURL=`${this.context.pageContext.site.absoluteUrl}/SiteAssets/CustomCSS.css`;
    SPComponentLoader.loadCss(cssURL);

    if (this.properties.FormType == "New") {
      const element: React.ReactElement<IPromotionResponseProps> = React.createElement(
        PromotionResponseNew,
        {          
          context: this.context,
          httpRequest:this.properties.httpRequest
        }
      );

      ReactDom.render(element, this.domElement);
    }
    else if (this.properties.FormType == "Edit") {
      const element: React.ReactElement<IPromotionResponseProps> = React.createElement(
        PromotionResponseEdit,
        {
          context: this.context,
          httpRequest:this.properties.httpRequest
        }
      );

      ReactDom.render(element, this.domElement);
    }
    else if (this.properties.FormType == "View") {
      const element: React.ReactElement<IPromotionResponseProps> = React.createElement(
        PromotionResponse,
        {
          context: this.context,
          httpRequest:this.properties.httpRequest
        }
      );

      ReactDom.render(element, this.domElement);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('FormType', {
                  label: 'Form Type',
                  options: [
                    { key: 'New', text: 'New' },
                    { key: 'Edit', text: 'Edit' },
                    { key: 'view', text: 'View' }
                  ]
                }),
                PropertyPaneTextField('httpRequest', {
                  label: "httpRequest"
                })

              ]
            }
          ]
        }
      ]
    };
  }
}
