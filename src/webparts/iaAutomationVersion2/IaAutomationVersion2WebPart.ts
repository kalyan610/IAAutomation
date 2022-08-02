import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'IaAutomationVersion2WebPartStrings';
import IaAutomationVersion2 from './components/IaAutomationVersion2';
import { IIaAutomationVersion2Props } from './components/IIaAutomationVersion2Props';

export interface IIaAutomationVersion2WebPartProps {
  description: string;
  RedirectURl:string;

}

export default class IaAutomationVersion2WebPart extends BaseClientSideWebPart<IIaAutomationVersion2WebPartProps> {

  public render(): void {
    const element: React.ReactElement<IIaAutomationVersion2Props> = React.createElement(
      IaAutomationVersion2,
      {
         url: "https://capcoinc.sharepoint.com/sites/capcointernalapplications/",
        //url: "https://capcoinc.sharepoint.com/sites/IAAutomationEnvironment/",
         context:this.context
        
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

                PropertyPaneTextField('RedirectURl', {
                  label: 'Enter url'
                })


              ]
            }
          ]
        }
      ]
    };
  }
}
