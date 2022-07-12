import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import ReadReceiptWebpart from './components/ReadReceiptWebpart';
import { IReadReceiptWebpartProps } from './components/IReadReceiptWebpartProps';
import { sp } from '@pnp/sp/presets/all';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import { ThemeProvider, ThemeChangedEventArgs  } from '@microsoft/sp-component-base';
import { __makeTemplateObject } from 'tslib';
import * as strings from 'ReadReceiptWebpartWebPartStrings';

export interface IReadReceiptWebpartWebPartProps {
  documentTitle: string; // description: string;
  storageList: string;
  acknowledgementLabel: string;
  acknowledgementMessage: string;
  readMessage: string;
}

export default class ReadReceiptWebpartWebPart extends BaseClientSideWebPart<IReadReceiptWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  public render(): void {
    const element: React.ReactElement<IReadReceiptWebpartProps> = React.createElement(
      ReadReceiptWebpart,
      {
        // description: this.properties.description,
        // isDarkTheme: this._isDarkTheme,
        // environmentMessage: this._environmentMessage,
        // hasTeamsContext: !!this.context.sdks.microsoftTeams,
        // userDisplayName: this.context.pageContext.user.displayName

        documentTitle: this.properties.documentTitle,
        currentUserDisplayName: this.context.pageContext.user.displayName,
        storageList: this.properties.storageList,
        acknowledgementLabel: this.properties.acknowledgementLabel,
        acknowledgementMessage: this.properties.acknowledgementMessage,
        readMessage: this.properties.readMessage,
        themeVariant: this._themeVariant,
        configured: this.properties.storageList ? this.properties.storageList !== '' : false,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // eslint-disable-next-line @microsoft/spfx/no-async-await
  protected async onInit(): Promise<void> {
   // this._environmentMessage = this._getEnvironmentMessage();

    //return super.onInit();

    await super.onInit();

    sp.setup(this.context);

    this._themeProvider = this.context.serviceScope.consume(
      ThemeProvider.serviceKey
    );

    this._themeVariant = this._themeProvider.tryGetTheme();

    this._themeProvider.themeChangedEvent.add(
      this, this._handleThemeChangedEvent
    );
  }

  private _handleThemeChangedEvent(args: ThemeChangedArgs): any {
    this._themeVariant = args.theme;
    this.render();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
                PropertyFieldListPicker('storageList', {
                  label: strings.StorageListLabel,
                  selectedList: this.properties.storageList,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key:'listPickerFieldId',
                  multiSelect: false,
                  baseTemplate: 100
                }),
                PropertyPaneTextField('documentTitle', {
                  label: strings.DocumentTitleLabel
                }),
                PropertyPaneTextField('acknowledgementLabel', {
                  label: strings.AcknowledgementLabelLabel
                }),
                PropertyPaneTextField('acknowledgementMessage', {
                  label: strings.AcknowledgementMessageLabel
                }),
                PropertyPaneTextField('readMessage', {
                  label: strings.ReadMessageLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
