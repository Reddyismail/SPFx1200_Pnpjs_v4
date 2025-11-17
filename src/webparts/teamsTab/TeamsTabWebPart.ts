import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TeamsTabWebPartStrings';
import TeamsTab from './components/TeamsTab';
import { ITeamsTabProps } from './components/ITeamsTabProps';
import { SPFx,spfi } from '@pnp/sp';
export interface ITeamsTabWebPartProps {
  description: string;
}

export default class TeamsTabWebPart extends BaseClientSideWebPart<ITeamsTabWebPartProps> {

  // private _isDarkTheme: boolean = false;
  // private _environmentMessage: string = '';
private _sp:any
  public render(): void {
    const element: React.ReactElement<ITeamsTabProps> = React.createElement(
      TeamsTab,
      {
        
        // description: this.properties.description,
        // isDarkTheme: this._isDarkTheme,
        // environmentMessage: this._environmentMessage,
        // hasTeamsContext: !!this.context.sdks.microsoftTeams,
        // userDisplayName: this.context.pageContext.user.displayName,
       sp:this._sp
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context)); // initialize PnPjs with SPFx context
    return super.onInit();
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
