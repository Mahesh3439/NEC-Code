import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown 
} from '@microsoft/sp-property-pane';

import * as strings from 'ProjectSummaryWebPartStrings';

import ProjectSummary from './components/ProjectSummary';
import ProjectSummarySubmit from './components/ProjectSummarySubmit';
import ProjectSummaryUpdate from './components/ProjectSummaryUpdate';
import ProjectSpace from './components/ProjectSpace';
import SummaryDetails from './components/SummaryDetails';
import ProjectApprovals from './components/ProjectApprovals';
import { IProjectSummaryProps } from './components/IProjectSummaryProps';

import * as CustomJS from 'CustomJS';

require('MultiFile');

export interface IProjectSummaryWebPartProps {
  description: string;
  FormType: string;
  context: WebPartContext;
 
}

export default class ProjectSummaryWebPart extends BaseClientSideWebPart<IProjectSummaryWebPartProps> {
  

  public render(): void {
    //const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
      let cssURL=`${this.context.pageContext.site.absoluteUrl}/SiteAssets/CustomCSS.css`;
      SPComponentLoader.loadCss(cssURL);
   
      if (this.properties.FormType == "New") {
        const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
          ProjectSummarySubmit,
          {          
            context: this.context
          }
        );
  
        ReactDom.render(element, this.domElement);
      }
      else if (this.properties.FormType == "Edit") {
        const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
          ProjectSummaryUpdate,
          {
            context: this.context
          }
        );
  
        ReactDom.render(element, this.domElement);
      }
      else if (this.properties.FormType == "ProjectSpace") {
        const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
          ProjectSpace,
          {
            context: this.context
          }
        );
  
        ReactDom.render(element, this.domElement);
      }
      else if (this.properties.FormType == "View") {
        const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
          SummaryDetails,
          {
            context: this.context
          }
        );
  
        ReactDom.render(element, this.domElement);
      }
      // else if (this.properties.FormType == "Approvals") {
      //   const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
      //     ProjectApprovals,
      //     {
      //       context: this.context,
      //       onDissmissPanel:()=>void
      //     }
      //   );
  
      //   ReactDom.render(element, this.domElement);
      // }
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
                    { key: 'Edit',text: 'Edit' },
                    { key: 'View', text: 'View' },
                    { key: 'ProjectSpace', text: 'ProjectSpace' }
                    // { key: 'Approvals', text: 'Approvals' }
                   
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
