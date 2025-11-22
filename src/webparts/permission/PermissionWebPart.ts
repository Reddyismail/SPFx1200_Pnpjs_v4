import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PermissionWebPartStrings';
import Permission from './components/Permission';
import { IPermissionProps } from './components/IPermissionProps';
import { spfi,SPFx } from '@pnp/sp';

export interface IPermissionWebPartProps {
  description: string;
}

export default class PermissionWebPart extends BaseClientSideWebPart<IPermissionWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp: any

  public render(): void {
    const element: React.ReactElement<IPermissionProps> = React.createElement(
      Permission,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        url: this.context.pageContext.web.absoluteUrl,
        image:this.context.pageContext.web.logoUrl,
        absoluteUrl: this.context.pageContext.web.absoluteUrl,
        permissionContext: this.context.pageContext.web.permissions,
        sp:this._sp,
        context: this.context


      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
     this._sp = spfi().using(SPFx(this.context)); // initialize PnPjs with SPFx context
     return super.onInit();
   }

 

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
