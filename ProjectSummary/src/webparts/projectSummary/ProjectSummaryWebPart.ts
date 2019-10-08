import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField 
} from '@microsoft/sp-property-pane';

import * as strings from 'ProjectSummaryWebPartStrings';

import ProjectSummary from './components/ProjectSummary';
import { IProjectSummaryProps } from './components/IProjectSummaryProps';
import  HTMLContent from './components/DevTest';
import {IApprovalsProps} from './components/DevTest';

import * as CustomJS from 'CustomJS';

require('MultiFile');

export interface IProjectSummaryWebPartProps {
  description: string;
  context: WebPartContext
 
}

export default class ProjectSummaryWebPart extends BaseClientSideWebPart<IProjectSummaryWebPartProps> {


  public render(): void {
    //const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
   
      const element: React.ReactElement<IProjectSummaryProps> = React.createElement(
       ProjectSummary,       
        {         
          context: this.context
        }
      );
      ReactDom.render(element, this.domElement);   
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
