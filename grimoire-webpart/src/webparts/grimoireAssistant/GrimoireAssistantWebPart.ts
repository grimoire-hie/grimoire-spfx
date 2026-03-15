import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneFieldType,
  type IPropertyPaneCustomFieldProps
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'GrimoireAssistantWebPartStrings';
import { GrimoireAssistant } from './components/GrimoireAssistant';
import { IGrimoireAssistantProps } from './components/IGrimoireAssistantProps';
import type { ISpThemeColors } from './store/useGrimoireStore';
import { DEFAULT_SP_THEME_COLORS } from './store/useGrimoireStore';
import {
  detectHostEnvironment,
  getHostChromeHeight
} from './utils/hostEnvironment';
import { getSP } from './services/pnp/pnpConfig';
import { PropertyPanePasswordField } from './propertyPane/PasswordField';
import {
  DEFAULT_PROXY_BACKEND,
  getProxyBackendOptions
} from './config/webPartDefaults';

export interface IGrimoireAssistantWebPartProps {
  description: string;
  proxyUrl?: string;
  proxyApiKey?: string;
  backendApiResource?: string;
  proxyBackend?: string;
  deploymentPrefix?: string;
  mcpEnvironmentId?: string;
}

export default class GrimoireAssistantWebPart extends BaseClientSideWebPart<IGrimoireAssistantWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _spThemeColors: ISpThemeColors = DEFAULT_SP_THEME_COLORS;
  private _resizeObserver: ResizeObserver | undefined;
  private _fallbackResizeHandler: (() => void) | undefined;

  public render(): void {
    this._updateAvailableHeight();

    const element: React.ReactElement<IGrimoireAssistantProps> = React.createElement(
      GrimoireAssistant,
      {
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        context: this.context,
        proxyUrl: this.properties.proxyUrl || '',
        proxyApiKey: this.properties.proxyApiKey || '',
        backendApiResource: this.properties.backendApiResource || '',
        proxyBackend: this.properties.proxyBackend || DEFAULT_PROXY_BACKEND,
        deploymentPrefix: this.properties.deploymentPrefix || '',
        mcpEnvironmentId: this.properties.mcpEnvironmentId || '',
        spThemeColors: this._spThemeColors
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    getSP(this.context);
    this._setupResizeObserver();
  }

  private _updateAvailableHeight(): void {
    const hasTeamsContext = !!this.context.sdks.microsoftTeams;
    const environment = detectHostEnvironment(hasTeamsContext);
    const chromeHeight = getHostChromeHeight(environment);
    const availableHeight = window.innerHeight - chromeHeight;

    this.domElement.style.setProperty('--grimoire-available-height', `${availableHeight}px`);
    this.domElement.style.setProperty('--grimoire-sp-chrome-height', `${chromeHeight}px`);
  }

  private _setupResizeObserver(): void {
    if (typeof ResizeObserver !== 'undefined') {
      this._resizeObserver = new ResizeObserver(() => {
        this._updateAvailableHeight();
      });
      this._resizeObserver.observe(document.body);
    } else {
      this._fallbackResizeHandler = () => this._updateAvailableHeight();
      window.addEventListener('resize', this._fallbackResizeHandler);
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);

      this._spThemeColors = {
        bodyBackground: semanticColors.bodyBackground || DEFAULT_SP_THEME_COLORS.bodyBackground,
        bodyText: semanticColors.bodyText || DEFAULT_SP_THEME_COLORS.bodyText,
        bodySubtext: semanticColors.bodySubtext || DEFAULT_SP_THEME_COLORS.bodySubtext,
        cardBackground: semanticColors.cardStandoutBackground || semanticColors.bodyBackground || DEFAULT_SP_THEME_COLORS.cardBackground,
        cardBorder: semanticColors.variantBorder || DEFAULT_SP_THEME_COLORS.cardBorder,
        isDark: this._isDarkTheme
      };
    }
  }

  protected onDispose(): void {
    if (this._resizeObserver) {
      this._resizeObserver.disconnect();
    }
    if (this._fallbackResizeHandler) {
      window.removeEventListener('resize', this._fallbackResizeHandler);
      this._fallbackResizeHandler = undefined;
    }
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
            },
            {
              groupName: strings.BackendGroupName,
              groupFields: [
                PropertyPaneTextField('proxyUrl', {
                  label: strings.ProxyUrlFieldLabel,
                  value: this.properties.proxyUrl || '',
                  description: strings.ProxyUrlFieldDescription
                }),
                this.createPasswordField('proxyApiKey', strings.ProxyApiKeyFieldLabel, strings.ProxyApiKeyFieldDescription),
                PropertyPaneTextField('backendApiResource', {
                  label: strings.BackendApiResourceFieldLabel,
                  value: this.properties.backendApiResource || '',
                  description: strings.BackendApiResourceFieldDescription
                }),
                PropertyPaneDropdown('proxyBackend', {
                  label: strings.ProxyBackendFieldLabel,
                  selectedKey: this.properties.proxyBackend || DEFAULT_PROXY_BACKEND,
                  options: getProxyBackendOptions({
                    reasoning: strings.ProxyBackendReasoningOptionLabel,
                    fast: strings.ProxyBackendFastOptionLabel
                  })
                }),
                PropertyPaneTextField('deploymentPrefix', {
                  label: strings.DeploymentPrefixFieldLabel,
                  value: this.properties.deploymentPrefix || '',
                  description: strings.DeploymentPrefixFieldDescription
                })
              ]
            },
            {
              groupName: strings.M365McpGroupName,
              groupFields: [
                PropertyPaneTextField('mcpEnvironmentId', {
                  label: strings.McpEnvironmentIdFieldLabel,
                  value: this.properties.mcpEnvironmentId || '',
                  description: strings.McpEnvironmentIdFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private getStringPropertyValue(targetProperty: keyof IGrimoireAssistantWebPartProps): string {
    const value = this.properties[targetProperty];
    return typeof value === 'string' ? value : '';
  }

  private setStringPropertyValue(targetProperty: keyof IGrimoireAssistantWebPartProps, newValue: string): string {
    const oldValue = this.getStringPropertyValue(targetProperty);
    this.properties[targetProperty] = newValue;
    return oldValue;
  }

  private createPasswordField(
    targetProperty: keyof IGrimoireAssistantWebPartProps,
    label: string,
    description?: string
  ): IPropertyPaneField<IPropertyPaneCustomFieldProps> {
    const onRender = (
      domElement: HTMLElement,
      _context: unknown,
      changeCallback?: (targetProp?: string, newValue?: string) => void
    ): void => {
      const value = this.getStringPropertyValue(targetProperty);
      const handleChange = (newValue: string): void => {
        const oldValue = this.setStringPropertyValue(targetProperty, newValue);
        this.onPropertyPaneFieldChanged(targetProperty as string, oldValue, newValue);
        if (changeCallback) {
          changeCallback(targetProperty as string, newValue);
        }
      };

      ReactDom.render(
        React.createElement(PropertyPanePasswordField, {
          label,
          description,
          value,
          onChange: handleChange
        }),
        domElement
      );
    };

    const onDispose = (domElement: HTMLElement): void => {
      ReactDom.unmountComponentAtNode(domElement);
    };

    return {
      type: PropertyPaneFieldType.Custom,
      targetProperty: targetProperty as string,
      properties: {
        key: `${String(targetProperty)}_password`,
        onRender,
        onDispose
      }
    } as IPropertyPaneField<IPropertyPaneCustomFieldProps>;
  }
}
