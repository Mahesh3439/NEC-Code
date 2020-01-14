import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'RndWebpartWebPartStrings';
import RndWebpart from './components/RndWebpart';
import { IRndWebpartProps } from './components/IRndWebpartProps';

import { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } from '@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor';

export interface IRndWebpartWebPartProps {
  description: string;
  htmlCode: string;
}

export default class RndWebpartWebPart extends BaseClientSideWebPart <IRndWebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRndWebpartProps> = React.createElement(
      RndWebpart,
      {
        description: this.properties.description,
        htmlCode:this.properties.htmlCode
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
                }),
                PropertyFieldCodeEditor('htmlCode', {
                  label: 'Edit HTML Code',
                  panelTitle: 'Edit HTML Code',
                  initialValue: this.properties.htmlCode,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  key: 'codeEditorFieldId',
                  language: PropertyFieldCodeEditorLanguages.HTML
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
