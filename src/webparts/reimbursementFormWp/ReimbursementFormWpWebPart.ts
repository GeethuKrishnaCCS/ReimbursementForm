import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ReimbursementFormWpWebPartStrings';
import ReimbursementFormWp from './components/ReimbursementFormWp';
import { IReimbursementFormWpProps, IReimbursementFormWpWebPartProps } from './interfaces/IReimbursementFormWpProps';



export default class ReimbursementFormWpWebPart extends BaseClientSideWebPart<IReimbursementFormWpWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IReimbursementFormWpProps> = React.createElement(
      ReimbursementFormWp,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        ClientList: this.properties.ClientList,
        ProgramList: this.properties.ProgramList,
        ProjectList: this.properties.ProjectList,
        AdminApprover: this.properties.AdminApprover,
        CategoryList: this.properties.CategoryList,
        DepartmentsList: this.properties.DepartmentsList,
        ReimbursementRequestList: this.properties.ReimbursementRequestList,
        ReimbursementItemsList: this.properties.ReimbursementItemsList,
        ReimbursementRequestSettingsList: this.properties.ReimbursementRequestSettingsList,
        SubcategoryList: this.properties.SubcategoryList,
        TasksList: this.properties.TasksList,

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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
                }),
                PropertyPaneTextField('AdminApprover', {
                  label: strings.AdminApprover
                }),
                PropertyPaneTextField('CategoryList', {
                  label: strings.CategoryList
                }),
                PropertyPaneTextField('ClientList', {
                  label: strings.ClientList
                }),
                PropertyPaneTextField('DepartmentsList', {
                  label: strings.DepartmentsList
                }),
                PropertyPaneTextField('ProgramList', {
                  label: strings.ProgramList
                }),
                PropertyPaneTextField('ProjectList', {
                  label: strings.ProjectList
                }),
                PropertyPaneTextField('ReimbursementRequestList', {
                  label: strings.ReimbursementRequestList
                }),
                PropertyPaneTextField('ReimbursementItemsList', {
                  label: strings.ReimbursementItemsList
                }),
                PropertyPaneTextField('ReimbursementRequestSettingsList', {
                  label: strings.ReimbursementRequestSettingsList
                }),
                PropertyPaneTextField('SubcategoryList', {
                  label: strings.SubcategoryList
                }),
                PropertyPaneTextField('TasksList', {
                  label: strings.TasksList
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}


