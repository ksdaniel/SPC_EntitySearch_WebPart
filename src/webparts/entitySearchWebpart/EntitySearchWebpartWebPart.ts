import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  type IPropertyPaneDropdownOption,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'EntitySearchWebpartWebPartStrings';
import EntitySearchWebpart from './components/EntitySearchWebpart';
import { IEntitySearchWebpartProps } from './components/IEntitySearchWebpartProps';

export interface IEntitySearchWebpartWebPartProps {
  description: string;
  listId: string;
  titleFieldInternalName: string;
  typeFieldInternalName: string;
  dealFieldInternalName: string;
  statusFieldInternalName: string;
}

export default class EntitySearchWebpartWebPart extends BaseClientSideWebPart<IEntitySearchWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _fieldOptions: IPropertyPaneDropdownOption[] = [];
  private _isLoadingLists: boolean = false;
  private _isLoadingFields: boolean = false;

  public render(): void {
    const element: React.ReactElement<IEntitySearchWebpartProps> = React.createElement(
      EntitySearchWebpart,
      {
        description: this.properties.description,
        listId: this.properties.listId,
        titleFieldInternalName: this.properties.titleFieldInternalName,
        typeFieldInternalName: this.properties.typeFieldInternalName,
        dealFieldInternalName: this.properties.dealFieldInternalName,
        statusFieldInternalName: this.properties.statusFieldInternalName,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._ensureDefaultFieldMappings();

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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this._loadListOptions();
    if (this.properties.listId) {
      await this._loadFieldOptions(this.properties.listId);
    }
    this.context.propertyPane.refresh();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listId' && oldValue !== newValue) {
      this._fieldOptions = [];
      this._resetFieldMappings();
      this.context.propertyPane.refresh();

      if (typeof newValue === 'string' && newValue) {
        await this._loadFieldOptions(newValue);
      }

      this.context.propertyPane.refresh();
      this.render();
    }
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
                PropertyPaneDropdown('listId', {
                  label: strings.ListFieldLabel,
                  options: this._listOptions,
                  disabled: this._isLoadingLists,
                  selectedKey: this.properties.listId
                }),
                PropertyPaneDropdown('titleFieldInternalName', {
                  label: strings.TitleMappingFieldLabel,
                  options: this._fieldOptions,
                  disabled: this._isLoadingFields || !this.properties.listId,
                  selectedKey: this.properties.titleFieldInternalName
                }),
                PropertyPaneDropdown('typeFieldInternalName', {
                  label: strings.TypeMappingFieldLabel,
                  options: this._fieldOptions,
                  disabled: this._isLoadingFields || !this.properties.listId,
                  selectedKey: this.properties.typeFieldInternalName
                }),
                PropertyPaneDropdown('dealFieldInternalName', {
                  label: strings.DealMappingFieldLabel,
                  options: this._fieldOptions,
                  disabled: this._isLoadingFields || !this.properties.listId,
                  selectedKey: this.properties.dealFieldInternalName
                }),
                PropertyPaneDropdown('statusFieldInternalName', {
                  label: strings.StatusMappingFieldLabel,
                  options: this._fieldOptions,
                  disabled: this._isLoadingFields || !this.properties.listId,
                  selectedKey: this.properties.statusFieldInternalName
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _ensureDefaultFieldMappings(): void {
    this.properties.titleFieldInternalName = this.properties.titleFieldInternalName || 'Title';
    this.properties.typeFieldInternalName = this.properties.typeFieldInternalName || 'EntityType';
    this.properties.dealFieldInternalName = this.properties.dealFieldInternalName || 'Deal';
    this.properties.statusFieldInternalName = this.properties.statusFieldInternalName || 'Status';
  }

  private _resetFieldMappings(): void {
    this.properties.titleFieldInternalName = '';
    this.properties.typeFieldInternalName = '';
    this.properties.dealFieldInternalName = '';
    this.properties.statusFieldInternalName = '';
  }

  private async _loadListOptions(): Promise<void> {
    this._isLoadingLists = true;
    this.context.propertyPane.refresh();

    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Id,Title,Hidden,BaseTemplate&$filter=Hidden eq false and BaseTemplate eq 100&$orderby=Title`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Unable to load lists: ${response.statusText}`);
      }

      const data = await response.json() as { value: Array<{ Id: string; Title: string }> };
      this._listOptions = data.value.map((list) => ({
        key: list.Id,
        text: list.Title
      }));
    } catch (error) {
      this._listOptions = [];
      console.error('EntitySearchWebPart: failed to load lists.', error);
    } finally {
      this._isLoadingLists = false;
    }
  }

  private async _loadFieldOptions(listId: string): Promise<void> {
    this._isLoadingFields = true;
    this.context.propertyPane.refresh();

    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,Title,Hidden,ReadOnlyField&$filter=Hidden eq false and ReadOnlyField eq false&$orderby=Title`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Unable to load fields: ${response.statusText}`);
      }

      const data = await response.json() as { value: Array<{ InternalName: string; Title: string }> };
      this._fieldOptions = data.value.map((field) => ({
        key: field.InternalName,
        text: `${field.Title} (${field.InternalName})`
      }));

      this._applyDefaultFieldMappingsFromOptions();
    } catch (error) {
      this._fieldOptions = [];
      console.error('EntitySearchWebPart: failed to load fields.', error);
    } finally {
      this._isLoadingFields = false;
    }
  }

  private _applyDefaultFieldMappingsFromOptions(): void {
    const availableKeys = new Set(this._fieldOptions.map(option => String(option.key)));

    this.properties.titleFieldInternalName = this._chooseFieldMapping(
      this.properties.titleFieldInternalName,
      ['Title'],
      availableKeys
    );
    this.properties.typeFieldInternalName = this._chooseFieldMapping(
      this.properties.typeFieldInternalName,
      ['EntityType', 'Type'],
      availableKeys
    );
    this.properties.dealFieldInternalName = this._chooseFieldMapping(
      this.properties.dealFieldInternalName,
      ['Deal'],
      availableKeys
    );
    this.properties.statusFieldInternalName = this._chooseFieldMapping(
      this.properties.statusFieldInternalName,
      ['Status'],
      availableKeys
    );
  }

  private _chooseFieldMapping(currentValue: string, fallbacks: string[], availableKeys: Set<string>): string {
    if (currentValue && availableKeys.has(currentValue)) {
      return currentValue;
    }

    const fallbackMatch = fallbacks.find(fallback => availableKeys.has(fallback));
    return fallbackMatch || '';
  }
}
