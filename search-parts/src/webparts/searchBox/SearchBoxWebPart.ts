import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, ServiceKey, Text, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { GlobalSettings } from 'office-ui-fabric-react';
import { IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { uniqBy } from '@microsoft/sp-lodash-subset';
import { DynamicProperty } from "@microsoft/sp-component-base";
import * as webPartStrings from 'SearchBoxWebPartStrings';
import {
    IPropertyPaneConfiguration,
    IPropertyPaneField,
    PropertyPaneSlider,
    PropertyPaneDropdown,
    PropertyPaneDynamicField,
    PropertyPaneDynamicFieldSet,
    PropertyPaneTextField,
    PropertyPaneToggle,
    DynamicDataSharedDepth,
    IPropertyPanePage,
    IPropertyPaneGroup,
    PropertyPaneLabel
} from "@microsoft/sp-property-pane";
import SearchBoxContainer from './components/SearchBoxContainer';
import { ISearchBoxContainerProps } from './components/ISearchBoxContainerProps';
import { DynamicDataService } from '../../services/dynamicDataService/DynamicDataService';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import IDynamicDataService from '../../services/dynamicDataService/IDynamicDataService';
import { ComponentType } from '../../common/ComponentType';
import { ISearchBoxWebPartProps } from './ISearchBoxWebPartProps';
import { UrlHelper, PageOpenBehavior, QueryPathBehavior } from '../../helpers/UrlHelper';
import * as commonStrings from 'CommonStrings';
import { ServiceScope } from '@microsoft/sp-core-library';
import { ISuggestionProviderDefinition, BaseSuggestionProvider } from '@pnp/modern-search-extensibility';
import { AvailableSuggestionProviders, BuiltinSuggestionProviderKeys } from '../../providers/AvailableSuggestionProviders';
import { ISuggestionProvider } from '@pnp/modern-search-extensibility';
import { ServiceScopeHelper } from '../../helpers/ServiceScopeHelper';
import { Toggle, IToggleProps, MessageBar, MessageBarType, Link } from "office-ui-fabric-react";
import { ISuggestionProviderConfiguration } from '../../providers/ISuggestionProviderConfiguration';
import { IExtensibilityConfiguration } from '../../models/common/IExtensibilityConfiguration';
import { Constants } from '../../common/Constants';
import { ITokenService } from '@pnp/modern-search-extensibility';
import { BuiltinTokenNames, TokenService } from '../../services/tokenService/TokenService';
import { BaseWebPart } from '../../common/BaseWebPart';
import { DynamicPropertyHelper } from '../../helpers/DynamicPropertyHelper';
import { INlpService } from '../../services/nlpService/INlpService';
import { NlpService } from '../../services/nlpService/NlpService';
import { ServiceHelper } from '../../helpers/ServiceHelper';

export default class SearchBoxWebPart extends BaseWebPart<ISearchBoxWebPartProps> implements IDynamicDataCallables {

    /**
     * The error message
     */
    private errorMessage: string = undefined;

    /**
     * Dynamically loaded components for property pane
     */
    private _propertyFieldCollectionData: any = null;
    private _propertyPanePropertyEditor = null;
    private _customCollectionFieldType: any = null;

    /**
     * The dynamic data service instance
     */
    private dynamicDataService: IDynamicDataService;

    /**
     * The search query text present in the search box
     */
    private _searchQueryText: string = '';

    /*
    * The service scope for this specific Web Part instance
    */
    private webPartInstanceServiceScope: ServiceScope;

    /**
     * The available custom suggestions providers
     */
    private availableCustomProviders: ISuggestionProviderDefinition[] = AvailableSuggestionProviders.BuiltinSuggestionProviders;

    /**
     * The current selected suggestion providers
     */
    private _selectedCustomProviders: ISuggestionProvider[] = [];

    private _pushStateCallback = null;

    /**
     * The token service instance
     */
    private tokenService: ITokenService;

    private nlpService: INlpService;

    private serviceHelper: ServiceHelper;

    constructor() {
        super();

        this._bindHashChange = this._bindHashChange.bind(this);
    }

    protected async onInit() {

        this.initializeProperties();

        // Initializes shared services
        await this.initializeBaseWebPart();

        // Initializes the Web Part instance services
        this.initializeWebPartServices();

        // Load extensibility libaries extensions
        await this.loadExtensions(this.wbProperties.extensibilityLibraryConfiguration);

        this._bindHashChange();
        this._handleQueryStringChange();

        this.context.dynamicDataSourceManager.initializeSource(this);

        return super.onInit();
    }

    public async render(): Promise<void> {

        try {

            // Reset the error message every time
            this.errorMessage = undefined;

            // Initialize provider instances
            this._selectedCustomProviders = await this.initializeSuggestionProviders(this.wbProperties.suggestionProviderConfiguration);

        } catch (error) {
            // Catch instanciation or wrong definition errors for extensibility scenarios
            this.errorMessage = error.message ? error.message : error;
        }

        if (this.context.propertyPane && this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
        }

        return this.renderCompleted();
    }

    protected renderCompleted(): void {

        if (!this.domElement) {
            return;
        }
        let renderRootElement: JSX.Element = null;

        let inputValue = "";
        if (this.wbProperties.queryText) {
            try {
                inputValue = DynamicPropertyHelper.tryGetValueSafe(this.wbProperties.queryText);
                if (inputValue !== undefined && typeof (inputValue) === 'string') {
                    inputValue = decodeURIComponent(inputValue);
                }

            } catch (error) {
                // Likely issue when q=%25 in spfx
            }
        }

        if (inputValue && typeof (inputValue) === 'string') {

            // Notify subscriber a new value if available
            this._searchQueryText = decodeURIComponent(inputValue);

            // Set the input query text globally for the page. There can be only one input query text submitted at a time even if multiple search box components are on the page
            GlobalSettings.setValue(BuiltinTokenNames.inputQueryText, this._searchQueryText);

            this.context.dynamicDataSourceManager.notifyPropertyChanged(ComponentType.SearchBox);
        }

        renderRootElement = React.createElement(SearchBoxContainer, {
            domElement: this.domElement,
            enableQuerySuggestions: this.wbProperties.enableQuerySuggestions,
            inputValue: this._searchQueryText,
            openBehavior: this.wbProperties.openBehavior,
            pageUrl: this.wbProperties.pageUrl,
            placeholderText: this.wbProperties.placeholderText,
            queryPathBehavior: this.wbProperties.queryPathBehavior,
            queryStringParameter: this.wbProperties.queryStringParameter,
            inputTemplate: this.wbProperties.inputTemplate,
            searchInNewPage: this.wbProperties.searchInNewPage,
            themeVariant: this._themeVariant,
            onSearch: this._onSearch,
            suggestionProviders: this._selectedCustomProviders,
            numberOfSuggestionsPerGroup: this.wbProperties.numberOfSuggestionsPerGroup,
            tokenService: this.tokenService,
            enableNlpService: this.wbProperties.enableNlpService,
            nlpService: this.nlpService,
            isStaging: this.wbProperties.isStaging,
            serviceScope: this.context.serviceScope
        } as ISearchBoxContainerProps);

        // Error message
        if (this.errorMessage) {
            renderRootElement = React.createElement(MessageBar, {
                messageBarType: MessageBarType.error,
            }, this.errorMessage, React.createElement(Link, {
                target: '_blank',
                href: this.wbProperties.documentationLink
            }, commonStrings.General.Resources.PleaseReferToDocumentationMessage));
        }

        ReactDom.render(renderRootElement, this.domElement);

        // This call set this.renderedOnce to 'true' so we need to execute it at the very end
        super.renderCompleted();
    }

    protected onDispose(): void {
        if (this._pushStateCallback) {
            window.history.pushState = this._pushStateCallback;
        }
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get isRenderAsync(): boolean {
        return true;
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

        let propertyPanePages: IPropertyPanePage[] = [];
        let providerOptionGroups: IPropertyPaneGroup[] = [];
        let extensibilityConfigurationGroups: IPropertyPaneGroup[] = [];


        if (this._selectedCustomProviders.length > 0 && !this.errorMessage) {
            this._selectedCustomProviders.forEach(provider => {
                providerOptionGroups = providerOptionGroups.concat(provider.getPropertyPaneGroupsConfiguration());
            });
        }

        propertyPanePages.push(
            {
                groups: [
                    {
                        groupName: webPartStrings.PropertyPane.SearchBoxSettingsGroup.GroupName,
                        groupFields: this._getSearchBoxSettingsFields()
                    },
                    {
                        groupName: webPartStrings.PropertyPane.SearchBoxQueryNlpSettingsGroup.GroupName,
                        groupFields: this._getSearchQueryOptimizationFields()
                    },
                ],
                displayGroupsAsAccordion: true
            },
            {
                groups: [
                    {
                        groupName: webPartStrings.PropertyPane.QuerySuggestionsGroup.GroupName,
                        groupFields: this._getSearchQuerySuggestionsFields()
                    },
                    ...providerOptionGroups
                ],
                displayGroupsAsAccordion: true
            },
            {
                groups: [
                    {
                        groupName: webPartStrings.PropertyPane.AvailableConnectionsGroup.GroupName,
                        groupFields: this._getSearchAvailableConnectionsFields()
                    }
                ],
                displayGroupsAsAccordion: true
            }
        );

        // Extensibility configuration
        extensibilityConfigurationGroups.push({
            groupName: commonStrings.PropertyPane.InformationPage.Extensibility.GroupName,
            groupFields: this.getExtensibilityFields()
        });


        // 'About' infos
        propertyPanePages.push(
            {
                displayGroupsAsAccordion: true,
                groups: [
                    ...this.getPropertyPaneWebPartInfoGroups(),
                    ...extensibilityConfigurationGroups,
                    {
                        groupName: commonStrings.PropertyPane.InformationPage.ImportExport,
                        groupFields: [this._propertyPanePropertyEditor({
                            webpart: this,
                            key: 'propertyEditor'
                        })]
                    }
                ]
            }
        );

        return {
            pages: propertyPanePages
        };
    }

    protected async loadPropertyPaneResources(): Promise<void> {

        const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
        );
        this._propertyFieldCollectionData = PropertyFieldCollectionData;
        this._customCollectionFieldType = CustomCollectionFieldType;

        const { PropertyPanePropertyEditor } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/PropertyPanePropertyEditor'
        );
        this._propertyPanePropertyEditor = PropertyPanePropertyEditor;
    }

    protected get propertiesMetadata(): IWebPartPropertiesMetadata {
        return {
            'queryText': {
                dynamicPropertyType: 'string'
            }
        };
    }

    protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {

        if (!this.wbProperties.useDynamicDataSource) {
            this.wbProperties.queryText.setValue('');
        }

        if (propertyPath.localeCompare('enableQuerySuggestions') === 0 && !newValue) {

            // Disable all providers
            this.wbProperties.suggestionProviderConfiguration.forEach(provider => {
                provider.enabled = false;
            });
        }

        if (propertyPath.localeCompare('extensibilityLibraryConfiguration') === 0) {

            // Remove duplicates if any
            const cleanConfiguration = uniqBy(this.wbProperties.extensibilityLibraryConfiguration, 'id');

            // Reset existing definitions to default
            this.availableCustomProviders = AvailableSuggestionProviders.BuiltinSuggestionProviders;

            await this.loadExtensions(cleanConfiguration);
        }

        this._bindHashChange();
    }

    protected async onPropertyPaneConfigurationStart() {
        await this.loadPropertyPaneResources();
    }

    public getPropertyDefinitions(): IDynamicDataPropertyDefinition[] {
        // Use the Web Part title as property title since we don't expose sub properties
        return [
            {
                id: ComponentType.SearchBox,
                title: webPartStrings.General.DynamicPropertyDefinition
            }
        ];
    }

    public getPropertyValue(propertyId: string) {
        switch (propertyId) {
            case ComponentType.SearchBox:
                return this._searchQueryText;
        }
    }

    /**
     * Initializes the query optimization data provider instance according to the current environment
     */
    private initNlpService() {
        this.serviceHelper = new ServiceHelper(this.context.httpClient);

        if (this.wbProperties.enableNlpService && this.wbProperties.nlpServiceUrl) {
            if (Environment.type === EnvironmentType.SharePoint) {
                this.nlpService = this.context.serviceScope.consume<INlpService>(NlpService.ServiceKey);
                this.nlpService.setServiceUrl(this.wbProperties.nlpServiceUrl);
            }
        }
    }

    /**
     * Determines the group fields for the search query options inside the property pane
     */
    private _getSearchAvailableConnectionsFields(): IPropertyPaneField<any>[] {

        // Sets up search query fields
        let searchAvailabeConnectionsConfigFields: IPropertyPaneField<any>[] = [
            PropertyPaneToggle('useDynamicDataSource', {
                label: webPartStrings.PropertyPane.AvailableConnectionsGroup.UseDynamicDataSourceLabel,
            })
        ];

        if (this.wbProperties.useDynamicDataSource) {
            searchAvailabeConnectionsConfigFields.push(
                PropertyPaneDynamicFieldSet({
                    label: webPartStrings.PropertyPane.AvailableConnectionsGroup.QueryKeywordsPropertyLabel,
                    fields: [
                        PropertyPaneDynamicField('queryText', {
                            label: webPartStrings.PropertyPane.AvailableConnectionsGroup.QueryKeywordsPropertyLabel,
                        })
                    ],
                    sharedConfiguration: {
                        depth: DynamicDataSharedDepth.Source,
                    }
                })
            );
        }

        return searchAvailabeConnectionsConfigFields;
    }

    private _getSearchQuerySuggestionsFields(): IPropertyPaneField<any>[] {

        let searchQuerySuggestionsFields: IPropertyPaneField<any>[] = [
            PropertyPaneToggle("enableQuerySuggestions", {
                label: webPartStrings.PropertyPane.QuerySuggestionsGroup.EnableQuerySuggestions
            })
        ];

        if (this.wbProperties.enableQuerySuggestions) {

            searchQuerySuggestionsFields.push(
                this._propertyFieldCollectionData('suggestionProviderConfiguration', {
                    manageBtnLabel: webPartStrings.PropertyPane.QuerySuggestionsGroup.EditSuggestionProvidersLabel,
                    key: 'suggestionProviderConfiguration',
                    panelHeader: webPartStrings.PropertyPane.QuerySuggestionsGroup.EditSuggestionProvidersLabel,
                    panelDescription: webPartStrings.PropertyPane.QuerySuggestionsGroup.SuggestionProvidersDescription,
                    disableItemCreation: true,
                    disableItemDeletion: true,
                    disabled: !this.wbProperties.enableQuerySuggestions,
                    label: webPartStrings.PropertyPane.QuerySuggestionsGroup.SuggestionProvidersLabel,
                    value: this.wbProperties.suggestionProviderConfiguration,
                    fields: [
                        {
                            id: 'enabled',
                            title: webPartStrings.PropertyPane.QuerySuggestionsGroup.EnabledPropertyLabel,
                            type: this._customCollectionFieldType.custom,
                            onCustomRender: (field, value, onUpdate, item, itemId) => {
                                return (
                                    React.createElement("div", null,
                                        React.createElement(Toggle, {
                                            key: itemId, checked: value, onChange: (evt, checked) => {
                                                onUpdate(field.id, checked);
                                            },
                                            offText: commonStrings.General.OffTextLabel,
                                            onText: commonStrings.General.OnTextLabel
                                        })
                                    )
                                );
                            }
                        },
                        {
                            id: 'name',
                            title: webPartStrings.PropertyPane.QuerySuggestionsGroup.ProviderNamePropertyLabel,
                            type: this._customCollectionFieldType.custom,
                            onCustomRender: (field, value) => {
                                return (
                                    React.createElement("div", { style: { 'fontWeight': 600 } }, value)
                                );
                            }
                        },
                        {
                            id: 'description',
                            title: webPartStrings.PropertyPane.QuerySuggestionsGroup.ProviderDescriptionPropertyLabel,
                            type: this._customCollectionFieldType.custom,
                            onCustomRender: (field, value) => {
                                return (
                                    React.createElement("div", null, value)
                                );
                            }
                        }
                    ]
                }),
                PropertyPaneSlider('numberOfSuggestionsPerGroup', {
                    min: 1,
                    max: 20,
                    showValue: true,
                    step: 1,
                    label: webPartStrings.PropertyPane.QuerySuggestionsGroup.NumberOfSuggestionsToShow
                })
            );
        }

        return searchQuerySuggestionsFields;
    }

    /**
     * Determines the group fields for the search options inside the property pane
     */
    private _getSearchBoxSettingsFields(): IPropertyPaneField<any>[] {

        let searchBehaviorOptionsFields: IPropertyPaneField<any>[] = [
            PropertyPaneTextField('placeholderText', {
                label: webPartStrings.PropertyPane.SearchBoxSettingsGroup.PlaceholderTextLabel
            }),
            PropertyPaneToggle("searchInNewPage", {
                label: webPartStrings.PropertyPane.SearchBoxSettingsGroup.SearchInNewPageLabel
            })
        ];


        if (this.wbProperties.searchInNewPage) {
            searchBehaviorOptionsFields = searchBehaviorOptionsFields.concat([
                PropertyPaneTextField('inputTemplate', {
                    label: webPartStrings.PropertyPane.SearchBoxSettingsGroup.QueryInputTransformationLabel,
                    multiline: true,
                    placeholder: `{${BuiltinTokenNames.inputQueryText}}`
                }),
                PropertyPaneTextField('pageUrl', {
                    disabled: !this.wbProperties.searchInNewPage,
                    label: webPartStrings.PropertyPane.SearchBoxSettingsGroup.PageUrlLabel,
                    onGetErrorMessage: this._validatePageUrl.bind(this),
                    validateOnFocusOut: true,
                    validateOnFocusIn: true,
                    placeholder: 'https://...'
                }),
                PropertyPaneDropdown('openBehavior', {
                    label: commonStrings.General.PageOpenBehaviorLabel,
                    options: [
                        { key: PageOpenBehavior.Self, text: commonStrings.General.SameTabOpenBehavior },
                        { key: PageOpenBehavior.NewTab, text: commonStrings.General.NewTabOpenBehavior }
                    ],
                    disabled: !this.wbProperties.searchInNewPage,
                    selectedKey: this.wbProperties.openBehavior
                }),
                PropertyPaneDropdown('queryPathBehavior', {
                    label: webPartStrings.PropertyPane.SearchBoxSettingsGroup.QueryPathBehaviorLabel,
                    options: [
                        { key: QueryPathBehavior.URLFragment, text: webPartStrings.PropertyPane.SearchBoxSettingsGroup.UrlFragmentQueryPathBehavior },
                        { key: QueryPathBehavior.QueryParameter, text: webPartStrings.PropertyPane.SearchBoxSettingsGroup.QueryStringQueryPathBehavior }
                    ],
                    disabled: !this.wbProperties.searchInNewPage,
                    selectedKey: this.wbProperties.queryPathBehavior
                })
            ]);
        }

        if (this.wbProperties.searchInNewPage && this.wbProperties.queryPathBehavior === QueryPathBehavior.QueryParameter) {
            searchBehaviorOptionsFields = searchBehaviorOptionsFields.concat([
                PropertyPaneTextField('queryStringParameter', {
                    disabled: !this.wbProperties.searchInNewPage || this.wbProperties.searchInNewPage && this.wbProperties.queryPathBehavior !== QueryPathBehavior.QueryParameter,
                    label: webPartStrings.PropertyPane.SearchBoxSettingsGroup.QueryStringParameterName,
                    onGetErrorMessage: (value) => {
                        if (this.wbProperties.queryPathBehavior === QueryPathBehavior.QueryParameter) {
                            if (value === null ||
                                value.trim().length === 0) {
                                return webPartStrings.PropertyPane.SearchBoxSettingsGroup.QueryParameterNotEmpty;
                            }
                        }
                        return '';
                    }
                })
            ]);
        }

        return searchBehaviorOptionsFields;
    }

    /**
     * Determines the group fields for the search query optimization inside the property pane
     */
    private _getSearchQueryOptimizationFields(): IPropertyPaneField<any>[] {

        let searchQueryOptimizationFields: IPropertyPaneField<any>[] = [
            PropertyPaneLabel("", {
                text: webPartStrings.PropertyPane.SearchBoxQueryNlpSettingsGroup.GroupDescription
            }),
            PropertyPaneToggle("enableNlpService", {
                checked: false,
                label: webPartStrings.PropertyPane.SearchBoxQueryNlpSettingsGroup.EnableNlpPropertyLabel,
            })
        ];

        if (this.wbProperties.enableNlpService) {

            searchQueryOptimizationFields.push(
                PropertyPaneTextField("nlpServiceUrl", {
                    label: webPartStrings.PropertyPane.SearchBoxQueryNlpSettingsGroup.ServiceUrlLabel,
                    disabled: !this.wbProperties.enableNlpService,
                    onGetErrorMessage: this._validateServiceUrl.bind(this),
                    description: Text.format(webPartStrings.PropertyPane.SearchBoxQueryNlpSettingsGroup.ServiceUrlDescription, window.location.host)
                }),
                PropertyPaneToggle("isStaging", {
                    label: webPartStrings.PropertyPane.SearchBoxQueryNlpSettingsGroup.UseStagingEndpoint,
                    disabled: !this.wbProperties.enableNlpService,
                }),
            );
        }

        return searchQueryOptimizationFields;
    }

    private getExtensibilityFields(): IPropertyPaneField<any>[] {

        let extensibilityFields: IPropertyPaneField<any>[] = [
            this._propertyFieldCollectionData('extensibilityLibraryConfiguration', {
                manageBtnLabel: commonStrings.PropertyPane.InformationPage.Extensibility.ManageBtnLabel,
                key: 'extensibilityLibraryConfiguration',
                enableSorting: true,
                panelHeader: webPartStrings.PropertyPane.InformationPage.Extensibility.PanelHeader,
                panelDescription: webPartStrings.PropertyPane.InformationPage.Extensibility.PanelDescription,
                label: commonStrings.PropertyPane.InformationPage.Extensibility.FieldLabel,
                value: this.wbProperties.extensibilityLibraryConfiguration,
                fields: [
                    {
                        id: 'name',
                        title: commonStrings.PropertyPane.InformationPage.Extensibility.Columns.Name,
                        type: this._customCollectionFieldType.string
                    },
                    {
                        id: 'id',
                        title: commonStrings.PropertyPane.InformationPage.Extensibility.Columns.Id,
                        type: this._customCollectionFieldType.string,
                        onGetErrorMessage: this._validateGuid.bind(this)
                    },
                    {
                        id: 'enabled',
                        title: commonStrings.PropertyPane.InformationPage.Extensibility.Columns.Enabled,
                        type: this._customCollectionFieldType.custom,
                        required: true,
                        onCustomRender: (field, value, onUpdate, item, itemId) => {
                            return (
                                React.createElement("div", null,
                                    React.createElement(Toggle, {
                                        key: itemId,
                                        checked: value,
                                        offText: commonStrings.General.OffTextLabel,
                                        onText: commonStrings.General.OnTextLabel,
                                        onChange: ((evt, checked) => {
                                            onUpdate(field.id, checked);
                                        }).bind(this)
                                    } as IToggleProps)
                                )
                            );
                        }
                    }
                ]
            })
        ];

        return extensibilityFields;
    }

    /**
     * Verifies if the string is a correct URL
     * @param value the URL to verify
     */
    private _validatePageUrl(value: string) {

        if ((!(/^(https?):\/\/[^\s/$.?#].[^\s]*/).test(value) || !value) && this.wbProperties.searchInNewPage) {
            return webPartStrings.PropertyPane.SearchBoxSettingsGroup.UrlErrorMessage;
        }

        return '';
    }

    /**
     * Initializes required Web Part wbProperties
     */
    private initializeProperties() {
        this.wbProperties.queryText = this.wbProperties.queryText ? this.wbProperties.queryText : new DynamicProperty<string>(this.context.dynamicDataProvider);
        this.wbProperties.inputTemplate = this.wbProperties.inputTemplate ? this.wbProperties.inputTemplate : `{${BuiltinTokenNames.inputQueryText}}`;

        this.wbProperties.openBehavior = this.wbProperties.openBehavior ? this.wbProperties.openBehavior : PageOpenBehavior.Self;
        this.wbProperties.queryPathBehavior = this.wbProperties.queryPathBehavior ? this.wbProperties.queryPathBehavior : QueryPathBehavior.URLFragment;

        this.wbProperties.suggestionProviderConfiguration = this.wbProperties.suggestionProviderConfiguration ? this.wbProperties.suggestionProviderConfiguration : [];
        this.wbProperties.numberOfSuggestionsPerGroup = this.wbProperties.numberOfSuggestionsPerGroup ? this.wbProperties.numberOfSuggestionsPerGroup : 10;

        this.wbProperties.providerProperties = this.wbProperties.providerProperties ? this.wbProperties.providerProperties : {};

        this.wbProperties.extensibilityLibraryConfiguration = this.wbProperties.extensibilityLibraryConfiguration ? this.wbProperties.extensibilityLibraryConfiguration : [{
            name: commonStrings.General.Extensibility.DefaultExtensibilityLibraryName,
            enabled: true,
            id: Constants.DEFAULT_EXTENSIBILITY_LIBRARY_COMPONENT_ID
        }];
    }

    private initializeWebPartServices(): void {
        this.initNlpService();
        this.tokenService = this.context.serviceScope.consume<ITokenService>(TokenService.ServiceKey);
        this.webPartInstanceServiceScope = this.context.serviceScope.startNewChild();
        this.dynamicDataService = this.webPartInstanceServiceScope.createDefaultAndProvide(DynamicDataService.ServiceKey);
        this.dynamicDataService.dynamicDataProvider = this.context.dynamicDataProvider;
        this.webPartInstanceServiceScope.finish();
    }

    /**
     * Handler used to notify data source subscribers when the input query is updated
     */
    private _onSearch = (searchQuery: string): void => {

        this._searchQueryText = searchQuery;

        // Set the input query text globally for the page. There can be only one input query text submitted at a time even if multiple search box components are on the page
        GlobalSettings.setValue(BuiltinTokenNames.inputQueryText, searchQuery);

        this.context.dynamicDataSourceManager.notifyPropertyChanged(ComponentType.SearchBox);
        this.context.dynamicDataSourceManager.notifySourceChanged();

        // Update URL with raw search query
        if (this.wbProperties.useDynamicDataSource && this.wbProperties.queryText && this.wbProperties.queryText.reference) {

            // this.wbProperties.defaultQueryKeywords.reference
            // "PageContext:UrlData:queryParameters.query"
            const refChunks = this.wbProperties.queryText.reference.split(':');

            if (refChunks.length >= 3) {
                const paramType = refChunks[2];

                if (paramType === 'fragment') {
                    window.history.pushState(undefined, undefined, `#${searchQuery}`);
                }
                else if (paramType.indexOf('queryParameters') !== -1) {
                    const paramChunks = paramType.split('.');
                    const queryTextParam = paramChunks.length === 2 ? paramChunks[1] : 'q';
                    const newUrl = UrlHelper.addOrReplaceQueryStringParam(window.location.href, queryTextParam, searchQuery);

                    if (window.location.href !== newUrl) {
                        window.history.pushState({ path: newUrl }, undefined, newUrl);
                    }
                }
            }
        }
    }

    /**
     * Subscribes to URL hash change if the dynamic property is set to the default 'URL Fragment' property
     */
    private _bindHashChange() {

        if (this.wbProperties.queryText.tryGetSource() && this.wbProperties.queryText.reference.localeCompare('PageContext:UrlData:fragment') === 0) {
            // Manually subscribe to hash change since the default property doesn't
            window.addEventListener('hashchange', this.render);
        } else {
            window.removeEventListener('hashchange', this.render);
        }
    }

    private async initializeSuggestionProviders(suggestionProviderConfiguration: ISuggestionProviderConfiguration[]): Promise<ISuggestionProvider[]> {

        const promises: Promise<ISuggestionProvider>[] = [];
        let selectedProviders: ISuggestionProvider[] = [];

        suggestionProviderConfiguration.forEach(configuration => {
            if (configuration.enabled) {
                promises.push(this.getSuggestionProviderInstance(configuration.key, this.availableCustomProviders));
            }
        });

        if (promises.length > 0) {
            selectedProviders = await Promise.all(promises);
        } else {
            selectedProviders = [];
        }

        return selectedProviders;
    }

    /**
     * Gets the suggestion provider instance according to the selected one
     * @param providerKey the selected suggestion provider provider key
     * @param suggestionProviderDefinitions the available source definitions
     * @returns the data source provider instance
     */
    private async getSuggestionProviderInstance(providerKey: string, suggestionProviderDefinitions: ISuggestionProviderDefinition[]): Promise<ISuggestionProvider> {

        let suggestionsProvider: ISuggestionProvider = undefined;
        let serviceKey: ServiceKey<ISuggestionProvider> = undefined;

        if (providerKey) {

            switch (providerKey) {

                // SharePoint Search static suggestions
                case BuiltinSuggestionProviderKeys.SharePointStaticSuggestions:

                    const { SharePointSuggestionProvider } = await import(
                        /* webpackChunkName: 'pnp-modern-search-sharepoint-static-suggestions' */
                        '../../providers/SharePointSuggestionProvider'
                    );

                    serviceKey = ServiceKey.create<ISuggestionProvider>('ModernSearchSharePointStaticSuggestionProvider', SharePointSuggestionProvider);
                    break;

                // Custom provider
                default:

                    // Gets the registered service key according to the selected provider definition 
                    const matchingDefinitions = suggestionProviderDefinitions.filter((provider) => { return provider.key === providerKey; });

                    // Can only have one data source instance per key
                    if (matchingDefinitions.length > 0) {
                        serviceKey = matchingDefinitions[0].serviceKey;
                    } else {
                        // Case when the extensibility library is removed from the catalog or the configuration
                        throw new Error(Text.format(commonStrings.General.Extensibility.ProviderDefinitionNotFound, providerKey));
                    }

                    break;
            }

            return new Promise<ISuggestionProvider>((resolve, reject) => {

                const childServiceScope = ServiceScopeHelper.registerChildServices(this.webPartInstanceServiceScope, [
                    serviceKey
                ]);

                childServiceScope.whenFinished(async () => {

                    suggestionsProvider = childServiceScope.consume<ISuggestionProvider>(serviceKey);

                    // Verifiy if the layout implements correctly the ILayout interface and BaseLayout methods
                    const isValidProvider = (providerInstance: ISuggestionProvider): providerInstance is BaseSuggestionProvider<any> => {
                        return (
                            (providerInstance as BaseSuggestionProvider<any>).getPropertyPaneGroupsConfiguration !== undefined &&
                            (providerInstance as BaseSuggestionProvider<any>).getSuggestions !== undefined &&
                            (providerInstance as BaseSuggestionProvider<any>).onPropertyUpdate !== undefined &&
                            (providerInstance as BaseSuggestionProvider<any>).onInit !== undefined
                        );
                    };

                    if (!isValidProvider(suggestionsProvider)) {
                        reject(new Error(Text.format(commonStrings.General.Extensibility.InvalidProviderInstance, providerKey)));
                    }

                    // Initialize the provider
                    if (suggestionsProvider) {

                        suggestionsProvider.properties = this.wbProperties.providerProperties;
                        suggestionsProvider.context = this.context;
                        await suggestionsProvider.onInit();

                        resolve(suggestionsProvider);
                    }
                });
            });
        }
    }

    /**
     * Loads extensions from registered extensibility librairies
     */
    private async loadExtensions(librariesConfiguration: IExtensibilityConfiguration[]) {

        const customSuggestionProviderConfiguration: ISuggestionProviderConfiguration[] = [];

        // Load extensibility library if present
        const extensibilityLibraries = await this.extensibilityService.loadExtensibilityLibraries(librariesConfiguration);

        // Load extensibility additions
        if (extensibilityLibraries.length > 0) {

            extensibilityLibraries.forEach(extensibilityLibrary => {
                // Add custom suggestions providers if any
                this.availableCustomProviders = this.availableCustomProviders.concat(extensibilityLibrary.getCustomSuggestionProviders());
            });
        }

        // Resolve the provider configuration for the property pane according to providers
        this.availableCustomProviders.forEach(provider => {

            if (!this.wbProperties.suggestionProviderConfiguration.some(p => p.key === provider.key)) {

                customSuggestionProviderConfiguration.push({
                    key: provider.key,
                    description: provider.description,
                    enabled: false,
                    name: provider.name
                });

            }
        });

        // Add custom providers to the available providers
        this.wbProperties.suggestionProviderConfiguration = this.wbProperties.suggestionProviderConfiguration.concat(customSuggestionProviderConfiguration);
    }

    /**
     * Ensures the string value is a valid GUID
     * @param value the result source id
     */
    private _validateGuid(value: string): string {
        if (value.length > 0) {
            if (!(/^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$/).test(value)) {
                return 'Invalid GUID';
            }
        }

        return '';
    }

    /**
     * Subscribes to URL query string change events using SharePoint page router
     */
    private _handleQueryStringChange() {

        // To avoid pushState modification from many components on the page (ex: search box, etc.), 
        // only subscribe to query string changes if the connected source is either the searc queyr or explicit query string parameter
        if (/^(PageContext:SearchData:searchQuery)|(PageContext:UrlData:queryParameters)/.test(this.wbProperties.queryText.reference)) {

            ((h) => {
                this._pushStateCallback = history.pushState;
                h.pushState = this.pushStateHandler.bind(this);
            })(window.history);
        }
    }

    private pushStateHandler(state, key, path) {

        this._pushStateCallback.apply(history, [state, key, path]);
        if (this.wbProperties.queryText.isDisposed) {
            return;
        }

        const source = this.wbProperties.queryText.tryGetSource();

        if (source && source.id === ComponentType.PageEnvironment) {
            this.render();
        }
    }

    /**
   * Ensures the service URL is valid 
   * @param value the service URL
   */
    private async _validateServiceUrl(value: string) {

        if ((!/^(https?):\/\/[^\s/$.?#].[^\s]*/.test(value) || !value)) {
            return webPartStrings.PropertyPane.SearchBoxQueryNlpSettingsGroup.ServiceUrlErrorMessage;
        } else {
            if (Environment.type !== EnvironmentType.Local) {
                try {
                    await this.serviceHelper.ensureUrlResovles(value);
                    return '';
                } catch (errorMessage) {
                    return Text.format(webPartStrings.PropertyPane.SearchBoxQueryNlpSettingsGroup.UrlNotResolvedErrorMessage, value, errorMessage);
                }
            } else {
                return '';
            }
        }
    }
}
