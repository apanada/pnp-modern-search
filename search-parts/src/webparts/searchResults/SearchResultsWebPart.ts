import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Text, DisplayMode, ServiceScope, Log, Guid } from '@microsoft/sp-core-library';
import { IComboBoxOption, Toggle, IToggleProps, MessageBarType, MessageBar, Link } from '@fluentui/react';
import { IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import * as webPartStrings from 'SearchResultsWebPartStrings';
import * as commonStrings from 'CommonStrings';
import { ISearchResultsContainerProps } from './components/ISearchResultsContainerProps';
import { IDataSource, IDataSourceDefinition, IComponentDefinition, ILayoutDefinition, ILayout, IDataFilter, LayoutType, FilterType, FilterComparisonOperator, BaseDataSource, IDataFilterValue, IDataFilterResult, FilterConditionOperator, IDataVertical } from '@pnp/modern-search-extensibility';
import {
    IPropertyPaneConfiguration,
    IPropertyPaneChoiceGroupOption,
    IPropertyPaneGroup,
    PropertyPaneChoiceGroup,
    IPropertyPaneField,
    PropertyPaneHorizontalRule,
    PropertyPaneToggle,
    PropertyPaneTextField,
    PropertyPaneSlider,
    IPropertyPanePage,
    PropertyPaneDropdown,
    PropertyPaneCheckbox,
    PropertyPaneDynamicField,
    DynamicDataSharedDepth,
    PropertyPaneDynamicFieldSet,
} from "@microsoft/sp-property-pane";
import ISearchResultsWebPartProps, { QueryTextSource } from './ISearchResultsWebPartProps';
import { AvailableDataSources, BuiltinDataSourceProviderKeys } from '../../dataSources/AvailableDataSources';
import { ServiceKey } from "@microsoft/sp-core-library";
import SearchResultsContainer from './components/SearchResultsContainer';
import { AvailableLayouts, BuiltinLayoutsKeys } from '../../layouts/AvailableLayouts';
import { ITemplateService } from '../../services/templateService/ITemplateService';
import { TemplateService } from '../../services/templateService/TemplateService';
import { ServiceScopeHelper } from '../../helpers/ServiceScopeHelper';
import { cloneDeep, flatten, isEmpty, isEqual, uniq, uniqBy } from "@microsoft/sp-lodash-subset";
import { AvailableComponents } from '../../components/AvailableComponents';
import { DynamicProperty } from '@microsoft/sp-component-base';
import { ITemplateSlot, IDataFilterToken, IDataFilterTokenValue, IDataContext, ITokenService } from '@pnp/modern-search-extensibility';
import { ResultTypeOperator } from '../../models/common/IDataResultType';
import { TokenService, BuiltinTokenNames } from '../../services/tokenService/TokenService';
import { TaxonomyService } from '../../services/taxonomyService/TaxonomyService';
import { SharePointSearchService } from '../../services/searchService/SharePointSearchService';
import IDynamicDataService from '../../services/dynamicDataService/IDynamicDataService';
import { IDataFilterSourceData } from '../../models/dynamicData/IDataFilterSourceData';
import { ComponentType, DynamicDataProperties } from '../../common/ComponentType';
import { DynamicDataService } from '../../services/dynamicDataService/DynamicDataService';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';
import { IDataResultSourceData } from '../../models/dynamicData/IDataResultSourceData';
import { LayoutHelper } from '../../helpers/LayoutHelper';
import { IAsyncComboProps } from '../../controls/PropertyPaneAsyncCombo/components/IAsyncComboProps';
import { AsyncCombo } from '../../controls/PropertyPaneAsyncCombo/components/AsyncCombo';
import { Constants } from '../../common/Constants';
import PnPTelemetry from "@pnp/telemetry-js";
import { IPageEventInfo } from '../../components/PaginationComponent';
import { DataFilterHelper } from '../../helpers/DataFilterHelper';
import { BuiltinFilterTemplates } from '../../layouts/AvailableTemplates';
import { IExtensibilityConfiguration } from '../../models/common/IExtensibilityConfiguration';
import { IDataVerticalSourceData } from '../../models/dynamicData/IDataVerticalSourceData';
import { BaseWebPart } from '../../common/BaseWebPart';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import commonStyles from '../../styles/Common.module.scss';
import { UrlHelper } from '../../helpers/UrlHelper';
import { ObjectHelper } from '../../helpers/ObjectHelper';
import { ItemSelectionMode } from '../../models/common/IItemSelectionProps';
import { PropertyPaneAsyncCombo } from '../../controls/PropertyPaneAsyncCombo/PropertyPaneAsyncCombo';
import { DynamicPropertyHelper } from '../../helpers/DynamicPropertyHelper';
import { spfi, SPFx } from '@pnp/sp';

const LogSource = "SearchResultsWebPart";

export default class SearchResultsWebPart extends BaseWebPart<ISearchResultsWebPartProps> implements IDynamicDataCallables {

    /**
     * The error message
     */
    private errorMessage: string = undefined;

    /**
     * Dynamic data related fields
     */
    private _filtersConnectionSourceData: DynamicProperty<IDataFilterSourceData>;
    private _verticalsConnectionSourceData: DynamicProperty<IDataVerticalSourceData>;

    private _currentDataResultsSourceData: IDataResultSourceData = {
        availableFieldsFromResults: [],
        availablefilters: [],
        selectedItems: []
    };

    /**
     * Dynamically loaded components for property pane
     */
    private _placeholderComponent: any = null;
    private _propertyFieldCodeEditor: any = null;
    private _propertyFieldCodeEditorLanguages: any = null;
    private _propertyFieldCollectionData: any = null;
    private _propertyFieldToogleWithCallout: any = null;
    private _propertyPaneWebPartInformation: any = null;
    private _propertyFieldCalloutTriggers: any = null;
    private _propertyFieldNumber: any = null;
    private _customCollectionFieldType: any = null;
    private _textDialogComponent: any = null;
    private _propertyPanePropertyEditor = null;

    /**
     * The selected data source for the WebPart
     */
    private dataSource: IDataSource;

    /**
     * Properties to avoid to recreate instances every render
     */
    private lastDataSourceKey: string;
    private lastLayoutKey: string;

    /**
     * The selected layout for the Web Part
     */
    private layout: ILayout;

    /**
     * The template content to display
     */
    private templateContentToDisplay: string;

    /**
     * The template service instance
     */
    private templateService: ITemplateService = undefined;

    /**
     * The available data source definitions (not instanciated)
     */
    private availableDataSourceDefinitions: IDataSourceDefinition[] = AvailableDataSources.BuiltinDataSources;

    /**
     * The available layout definitions (not instanciated)
     */
    private availableLayoutDefinitions: ILayoutDefinition[] = AvailableLayouts.BuiltinLayouts.filter(layout => { return layout.type === LayoutType.Results; });

    /**
     * The available web component definitions (not registered yet)
     */
    private availableWebComponentDefinitions: IComponentDefinition<any>[] = AvailableComponents.BuiltinComponents;

    /**
     * The current page number
     */
    private currentPageNumber: number = 1;

    /**
     * The page URL link if provided by the data source
     */
    private currentPageLinkUrl: string = null;

    /**
     * The available page links available in the pagination control
     */
    private availablePageLinks: string[] = [];

    /**
     * The token service instance
     */
    private tokenService: ITokenService;

    /**
     * the dynamic data service instance
     */
    private dynamicDataService: IDynamicDataService;

    /**
     * The service scope for this specific Web Part instance
     */
    private webPartInstanceServiceScope: ServiceScope;

    private _lastSelectedFilters: IDataFilter[] = [];
    private _lastInputQueryText: string = undefined;

    /**
     * The default template slots when the data source is instanciated for the first time
     */
    private _defaultTemplateSlots: ITemplateSlot[];

    private _pushStateCallback = null;

    /**
     * The available connections as property pane group
     */
    private propertyPaneConnectionsGroup: IPropertyPaneGroup[] = [];

    constructor() {
        super();

        this._bindHashChange = this._bindHashChange.bind(this);
        this._onDataRetrieved = this._onDataRetrieved.bind(this);
        this._onItemSelected = this._onItemSelected.bind(this);
    }

    public async render(): Promise<void> {

        // Determine the template content to display
        // In the case of an external template is selected, the render is done asynchronously waiting for the content to be fetched
        await this.initTemplate();

        // Refresh the token values with the latest information from environment (i.e connections and settings)
        await this.setTokens();

        // We resolve data source and layout instances directly in the render method to avoid unexpected render triggers due to Web Part property bag manipulation 
        // SPFx has an inner routine in reactive mode to trigger a render every time a property bag value is updated conflicting with the way data source and layouts share properties (see _afterPropertyUpdated)

        try {

            // Reset the error message every time
            this.errorMessage = undefined;

            // Get and initialize the data source instance if different (i.e avoid to create a new instance every time)
            if (this.lastDataSourceKey !== this.wbProperties.dataSourceKey) {
                this.dataSource = await this.getDataSourceInstance(this.wbProperties.dataSourceKey);
                this.lastDataSourceKey = this.wbProperties.dataSourceKey;
            }

            // Get and initialize layout instance if different (i.e avoid to create a new instance every time)
            if (this.lastLayoutKey !== this.wbProperties.selectedLayoutKey) {
                this.layout = await LayoutHelper.getLayoutInstance(this.webPartInstanceServiceScope, this.context, this.wbProperties, this.wbProperties.selectedLayoutKey, this.availableLayoutDefinitions);
                this.lastLayoutKey = this.wbProperties.selectedLayoutKey;
            }

        } catch (error) {
            // Catch instanciation or wrong definition errors for extensibility scenarios
            this.errorMessage = error.message ? error.message : error;
        }

        // Refresh the token values with the latest information from environment (i.e connections and settings)
        await this.setTokens();

        // Refresh the property pane to get layout and data source options
        if (this.context && this.context.propertyPane && this.context.propertyPane.isPropertyPaneOpen()) {
            this.context.propertyPane.refresh();
        }

        return this.renderCompleted();
    }

    public getPropertyDefinitions(): IDynamicDataPropertyDefinition[] {

        // Use the Web Part title as property title since we don't expose sub properties
        let propertyDefinitions: IDynamicDataPropertyDefinition[] = [];

        if (this.wbProperties.itemSelectionProps.allowItemSelection) {
            propertyDefinitions.push({
                id: DynamicDataProperties.AvailableFieldValuesFromResults,
                title: webPartStrings.PropertyPane.ConnectionsPage.AvailableFieldValuesFromResults,
            });
        }

        propertyDefinitions.push(
            {
                id: ComponentType.SearchResults,
                title: this.wbProperties.title ? `${this.wbProperties.title} - ${this.instanceId}` : `${webPartStrings.General.WebPartDefaultTitle} - ${this.instanceId}`,
            }
        );

        return propertyDefinitions;
    }

    public getPropertyValue(propertyId: string) {

        switch (propertyId) {
            case ComponentType.SearchResults:

                // Pass the Handlebars context to consumers, so they can register custom helpers for their own services 
                this._currentDataResultsSourceData.handlebarsContext = this.templateService.Handlebars;
                this._currentDataResultsSourceData.totalCount = this.dataSource?.getItemCount();

                return this._currentDataResultsSourceData;

            case DynamicDataProperties.AvailableFieldValuesFromResults:

                // Dynamic data values should be flatten https://docs.microsoft.com/en-us/sharepoint/dev/spfx/dynamic-data
                let fields = {};
                this._currentDataResultsSourceData.availableFieldsFromResults.forEach((field: string) => {

                    // Aggregate all values for this specific field across all items
                    // Ex:
                    // "FileType":['docx','pdf']
                    fields[field] = [];
                    this._currentDataResultsSourceData.selectedItems.forEach(selectedItem => {
                        const fieldValue = ObjectHelper.byPath(selectedItem, field);

                        // Special case where there value is a taxonomy item. In this case, we only take the GP0 part as it won't work otherwise with SharePoint search refiners or KQL conditions
                        const taxonomyItemRegExp = /GP0\|#0?((\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1})/gi;

                        if (taxonomyItemRegExp.test(fieldValue)) {
                            fieldValue.match(taxonomyItemRegExp).forEach(match => {
                                fields[field].push(match);
                            });
                        } else {

                            if (fieldValue) {
                                // Break down multiple values in a field value (like a multi choice or taxonomy column)
                                fieldValue.split(";").forEach(value => {
                                    fields[field].push(value);
                                });
                            } else {
                                fields[field].push(undefined);
                            }
                        }
                    });
                });

                return fields;
        }

        throw new Error('Bad property id');
    }

    protected renderCompleted(): void {

        let renderRootElement: JSX.Element = null;
        let renderDataContainer: JSX.Element = null;

        if (this.dataSource) {

            const dataContext = this.getDataContext();

            // The main content WP logic
            renderDataContainer = React.createElement(SearchResultsContainer, {
                dataSource: this.dataSource,
                dataSourceKey: this.wbProperties.dataSourceKey,
                templateContent: this.templateContentToDisplay,
                instanceId: this.instanceId,
                properties: JSON.parse(JSON.stringify(this.wbProperties)), // Create a copy to avoid unexpected reference value updates from data sources 
                onDataRetrieved: this._onDataRetrieved,
                onItemSelected: this._onItemSelected,
                pageContext: this.context.pageContext,
                dataContext: dataContext,
                themeVariant: this._themeVariant,
                serviceScope: this.webPartInstanceServiceScope,
                domElement: this.domElement,
                webPartTitleProps: {
                    displayMode: this.displayMode,
                    title: this.wbProperties.title,
                    updateProperty: (value: string) => {
                        this.wbProperties.title = value;
                    },
                    themeVariant: this._themeVariant,
                    className: commonStyles.wpTitle
                }
            } as ISearchResultsContainerProps);

            renderRootElement = renderDataContainer;

        } else {

            if (this.displayMode === DisplayMode.Edit) {
                const placeholder: React.ReactElement<any> = React.createElement(
                    this._placeholderComponent,
                    {
                        iconName: 'Database',
                        iconText: webPartStrings.General.PlaceHolder.IconText,
                        description: webPartStrings.General.PlaceHolder.Description,
                        buttonLabel: webPartStrings.General.PlaceHolder.ConfigureBtnLabel,
                        onConfigure: () => { this.context.propertyPane.openDetails(); }
                    }
                );
                renderRootElement = placeholder;
            } else {
                renderRootElement = null;
                Log.verbose(`[SearchResultsWebPart.renderCompleted]`, `The 'renderRootElement' was null during render.`, this.webPartInstanceServiceScope);
            }
        }

        // Check if the Web part is connected to a data vertical
        if (this._verticalsConnectionSourceData && this.wbProperties.selectedVerticalKeys.length > 0) {
            const verticalData = DynamicPropertyHelper.tryGetValueSafe(this._verticalsConnectionSourceData);

            // Remove the blank space introduced by the control zone when the Web Part displays nothing
            // WARNING: in theory, we are not supposed to touch DOM outside of the Web Part root element, This will break if the page attribute change
            const parentControlZone = this.getParentControlZone();

            // If the current selected vertical is not the one configured for this Web Part, we show nothing
            if (verticalData && this.wbProperties.selectedVerticalKeys.indexOf(verticalData.selectedVertical.key) === -1) {

                if (this.displayMode === DisplayMode.Edit) {

                    if (parentControlZone) {
                        parentControlZone.removeAttribute('style');
                    }

                    // Get tab name of selected verticals
                    const verticalNames = verticalData.verticalsConfiguration.filter(cfg => {
                        return this.wbProperties.selectedVerticalKeys.indexOf(cfg.key) !== -1;
                    }).map(v => v.tabName);

                    renderRootElement = React.createElement('div', {},
                        React.createElement(
                            MessageBar, {
                            messageBarType: MessageBarType.info,
                        },
                            Text.format(commonStrings.General.CurrentVerticalNotSelectedMessage, verticalNames.join(','))
                        ),
                        renderRootElement
                    );
                } else {
                    renderRootElement = null;

                    // Reset data source information
                    this._currentDataResultsSourceData = {
                        availableFieldsFromResults: [],
                        availablefilters: []
                    };

                    // Remove margin and padding for the empty control zone
                    if (parentControlZone) {
                        parentControlZone.setAttribute('style', 'margin-top:0px;padding:0px');
                    }

                }

            } else {

                if (parentControlZone) {
                    parentControlZone.removeAttribute('style');
                }
            }
        }

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

    protected async onInit(): Promise<void> {


        // Initializes Web Part properties
        this.initializeProperties();

        // Initializes shared services
        await this.initializeBaseWebPart();

        // Initializes the Web Part instance services
        this.initializeWebPartServices();

        // Bind web component events
        this.bindPagingEvents();

        this._bindHashChange();
        this._handleQueryStringChange();

        // Load extensibility libaries extensions
        await this.loadExtensions(this.wbProperties.extensibilityLibraryConfiguration);

        // Register Web Components in the global page context. We need to do this BEFORE the template processing to avoid race condition.
        // Web components are only defined once.
        await this.templateService.registerWebComponents(this.availableWebComponentDefinitions, this.instanceId);

        try {
            // Disable PnP Telemetry
            const telemetry = PnPTelemetry.getInstance();
            telemetry.optOut();
        } catch (error) {
            Log.warn(LogSource, `Opt out for PnP Telemetry failed. Details: ${error}`, this.context.serviceScope);
        }

        if (this.wbProperties.dataSourceKey && this.wbProperties.selectedLayoutKey && this.wbProperties.enableTelemetry) {

            const usageEvent = {
                name: Constants.PNP_MODERN_SEARCH_EVENT_NAME,
                properties: {
                    selectedDataSource: this.wbProperties.dataSourceKey,
                    selectedLayout: this.wbProperties.selectedLayoutKey,
                    useFilters: this.wbProperties.useFilters,
                    useVerticals: this.wbProperties.useVerticals
                }
            };

            // Track event with application insights (PnP)
            const appInsights = new ApplicationInsights({
                config: {
                    maxBatchInterval: 0,
                    instrumentationKey: Constants.PNP_APP_INSIGHTS_INSTRUMENTATION_KEY,
                    namePrefix: LogSource,
                    disableFetchTracking: true,
                    disableAjaxTracking: true
                }
            });

            appInsights.loadAppInsights();
            appInsights.context.application.ver = this.manifest.version;
            appInsights.trackEvent(usageEvent);
        }

        // Initializes MS Graph Toolkit
        if (this.wbProperties.useMicrosoftGraphToolkit) {
            await this.loadMsGraphToolkit();
        }

        // Initializes this component as a discoverable dynamic data source
        this.context.dynamicDataSourceManager.initializeSource(this);

        if (this.displayMode === DisplayMode.Edit) {
            const { Placeholder } = await import(
                /* webpackChunkName: 'pnp-modern-search-property-pane' */
                '@pnp/spfx-controls-react/lib/Placeholder'
            );
            this._placeholderComponent = Placeholder;
        }

        // Initializes dynamic data connections. This could trigger a render if a connection is made with an other component resulting to a render race condition.
        this.ensureDynamicDataSourcesConnection();

        await super.onInit();

        const sp = spfi().using(SPFx(this.context));
    }

    protected onDispose(): void {
        if (this._pushStateCallback) {
            window.history.pushState = this._pushStateCallback;
        }
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get propertiesMetadata(): IWebPartPropertiesMetadata {
        return {
            'filtersData': {
                dynamicPropertyType: 'object'
            },
            'queryText': {
                dynamicPropertyType: 'string'
            }
        };
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

        let propertyPanePages: IPropertyPanePage[] = [];
        let layoutSlotsGroup: IPropertyPaneGroup[] = [];
        let commonDataSourceProperties: IPropertyPaneGroup[] = [];
        let extensibilityConfigurationGroups: IPropertyPaneGroup[] = [];

        // Retrieve the property settings for the data source provider
        let dataSourceProperties: IPropertyPaneGroup[] = [];

        // Data source options page
        propertyPanePages.push(
            {
                groups: [
                    {
                        groupName: webPartStrings.PropertyPane.DataSourcePage.DataSourceConnectionGroupName,
                        groupFields: [
                            PropertyPaneChoiceGroup('dataSourceKey', {
                                options: this.getDataSourceOptions()
                            })
                        ]
                    }
                ],
                displayGroupsAsAccordion: true
            }
        );

        // A data source is selected
        if (this.dataSource && !this.errorMessage) {

            dataSourceProperties = this.dataSource.getPropertyPaneGroupsConfiguration();

            // Add template slots if any
            if (this.dataSource.getTemplateSlots().length > 0) {
                layoutSlotsGroup = [{
                    groupName: webPartStrings.PropertyPane.DataSourcePage.TemplateSlots.GroupName,
                    groupFields: this.getTemplateSlotOptions()
                }];
            }

            // Add data source options to the first page
            propertyPanePages[0].groups = propertyPanePages[0].groups.concat([
                ...layoutSlotsGroup,
                // Load data source specific properties
                ...dataSourceProperties,
                ...commonDataSourceProperties,
                {
                    groupName: webPartStrings.PropertyPane.DataSourcePage.PagingOptionsGroupName,
                    groupFields: this.getPagingGroupFields()
                }
            ]);

            // Other pages
            propertyPanePages.push(
                {
                    displayGroupsAsAccordion: true,
                    groups: this.getStylingPageGroups()
                },
                {
                    groups: [
                        ...this.propertyPaneConnectionsGroup
                    ],
                    displayGroupsAsAccordion: true
                }
            );
        }

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
                        groupFields: [
                            this._propertyPanePropertyEditor({
                                webpart: this,
                                key: 'propertyEditor'
                            }),
                            PropertyPaneToggle('enableTelemetry', {
                                label: webPartStrings.PropertyPane.InformationPage.EnableTelemetryLabel,
                                offText: webPartStrings.PropertyPane.InformationPage.EnableTelemetryOn,
                                onText: webPartStrings.PropertyPane.InformationPage.EnableTelemetryOff,
                            })
                        ]
                    }
                ]
            }
        );

        return {
            pages: propertyPanePages
        };
    }

    protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {

        // Bind connected data sources
        if (propertyPath.localeCompare('filtersDataSourceReference') === 0 && this.wbProperties.filtersDataSourceReference ||
            propertyPath.localeCompare('verticalsDataSourceReference') === 0 && this.wbProperties.verticalsDataSourceReference
        ) {
            this.ensureDynamicDataSourcesConnection();
            this.context.propertyPane.refresh();
        }

        if (propertyPath.localeCompare('useFilters') === 0) {
            if (!this.wbProperties.useFilters) {
                this.wbProperties.filtersDataSourceReference = undefined;
                this._filtersConnectionSourceData = undefined;
                this.context.dynamicDataSourceManager.notifyPropertyChanged(ComponentType.SearchResults);
            }
        }

        if (propertyPath.localeCompare('useVerticals') === 0) {
            if (!this.wbProperties.useVerticals) {
                this.wbProperties.verticalsDataSourceReference = undefined;
                this.wbProperties.selectedVerticalKeys = [];
                this._verticalsConnectionSourceData = undefined;
            }
        }

        if (propertyPath.localeCompare('useDynamicFiltering') === 0 && !this.wbProperties.useDynamicFiltering) {
            this.wbProperties.selectedItemFieldValue.setValue('');
            this.wbProperties.selectedItemFieldValue.unregister(this.render);
        }

        // Detect if the layout has been changed to custom
        if (propertyPath.localeCompare('inlineTemplateContent') === 0) {

            // Automatically switch the option to 'Custom' if a default template has been edited
            // (meaning the user started from a default template)
            if (this.wbProperties.inlineTemplateContent && this.wbProperties.selectedLayoutKey !== BuiltinLayoutsKeys.ResultsCustom) {
                this.wbProperties.selectedLayoutKey = BuiltinLayoutsKeys.ResultsCustom;

                // Reset also the template URL
                this.wbProperties.externalTemplateUrl = '';

                // Reset the layout options (otherwise we stay with the previous layout options)
                this.context.propertyPane.refresh();
            }
        }

        // Notify data source a property has been updated (only if the data source is already selected)
        if ((propertyPath.localeCompare('dataSourceKey') !== 0) && this.dataSource) {
            this.dataSource.onPropertyUpdate(propertyPath, oldValue, newValue);
        }

        // Reset the data source properties
        if (propertyPath.localeCompare('dataSourceKey') === 0 && !isEqual(oldValue, newValue)) {

            // Reset dynamic data source data
            this._currentDataResultsSourceData.availablefilters = [];
            this._currentDataResultsSourceData.availableFieldsFromResults = [];

            // Notfify dynamic data consumers data have changed
            this.context.dynamicDataSourceManager.notifyPropertyChanged(ComponentType.SearchResults);

            this.wbProperties.dataSourceProperties = {};
            this.wbProperties.templateSlots = null;

            // Reset paging information
            this.currentPageNumber = 1;

            this._resetPagingData();
        }

        // Reset layout properties
        if (propertyPath.localeCompare('selectedLayoutKey') === 0 && !isEqual(oldValue, newValue) && this.wbProperties.selectedLayoutKey !== BuiltinLayoutsKeys.ResultsDebug.toString()) {
            this.wbProperties.layoutProperties = {};
        }

        // Notify layout a property has been updated (only if the layout is already selected)
        if ((propertyPath.localeCompare('selectedLayoutKey') !== 0) && this.layout) {
            this.layout.onPropertyUpdate(propertyPath, oldValue, newValue);
        }

        // Remove the connection when static query text or unused
        if ((propertyPath.localeCompare('queryTextSource') === 0 && this.wbProperties.queryTextSource === QueryTextSource.StaticValue) ||
            (propertyPath.localeCompare('queryTextSource') === 0 && oldValue === QueryTextSource.StaticValue && newValue === QueryTextSource.DynamicValue) ||
            (propertyPath.localeCompare('useInputQueryText') === 0 && !this.wbProperties.useInputQueryText)) {

            if (this.wbProperties.queryText.tryGetSource()) {
                this.wbProperties.queryText.unregister(this.render);
            }

            this.wbProperties.queryText.setValue('');
        }

        // Update template slots when default slots from data source change (ex: OData client type)
        if (propertyPath.indexOf('dataSourceProperties') !== -1 && this.dataSource && this._defaultTemplateSlots && !isEqual(this._defaultTemplateSlots, this.dataSource.getTemplateSlots())) {
            this.wbProperties.templateSlots = this.dataSource.getTemplateSlots();
            this._defaultTemplateSlots = this.dataSource.getTemplateSlots();
        }

        if (propertyPath.localeCompare('paging.itemsCountPerPage') === 0) {
            this._resetPagingData();
        }

        if (propertyPath.localeCompare('extensibilityLibraryConfiguration') === 0) {

            // Remove duplicates if any
            const cleanConfiguration = uniqBy(this.wbProperties.extensibilityLibraryConfiguration, 'id');

            // Reset existing definitions to default
            this.availableDataSourceDefinitions = AvailableDataSources.BuiltinDataSources;
            this.availableLayoutDefinitions = AvailableLayouts.BuiltinLayouts.filter(layout => { return layout.type === LayoutType.Results; });
            this.availableWebComponentDefinitions = AvailableComponents.BuiltinComponents;

            await this.loadExtensions(cleanConfiguration);
        }

        if (this.wbProperties.queryTextSource === QueryTextSource.StaticValue || !this.wbProperties.useDefaultQueryText || !this.wbProperties.useInputQueryText) {
            // Reset the default query text
            this.wbProperties.defaultQueryText = undefined;
        }

        if (propertyPath.localeCompare("useMicrosoftGraphToolkit") === 0 && this.wbProperties.useMicrosoftGraphToolkit) {

            // We load this dynamically to avoid tokens renewal failure at page load and decrease the bundle size. Most of the time, MGT won't be used in templates.
            await this.loadMsGraphToolkit();
        }

        if (propertyPath.localeCompare('selectedItemFieldValue') === 0) {

            const reference = this.wbProperties.selectedItemFieldValue.reference;

            // Reset the default SPFx property pane field automatically as this configuration is not allowed for this scenario
            if (reference && reference.indexOf(ComponentType.SearchResults) !== -1) {
                this.wbProperties.selectedItemFieldValue.setValue('');
                this.wbProperties.selectedItemFieldValue.unregister(this.render);
            } else {
                if (!oldValue.reference) {
                    this.wbProperties.selectedItemFieldValue.register(this.render);
                }
            }
        }

        if (propertyPath.localeCompare('itemSelectionProps.destinationFieldName') === 0 && !isEqual(oldValue, newValue)) {

            const filterToken = this.tokenService.getTokenValue(BuiltinTokenNames.filters);

            if (filterToken) {
                // Reset previous token value 
                delete filterToken[oldValue];
            }

        }

        // Refresh list of available connections
        this.propertyPaneConnectionsGroup = await this.getConnectionOptionsGroup();
        this.context.propertyPane.refresh();

        // Reset the page number to 1 every time the Web Part properties change
        this.currentPageNumber = 1;
    }

    public onCustomPropertyUpdate(propertyPath: string, newValue: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {

        if (propertyPath.localeCompare('selectedVerticalKeys') === 0) {
            changeCallback(propertyPath, (cloneDeep(newValue) as IComboBoxOption[]).map(v => { return v.key as string; }));
            this.context.propertyPane.refresh();
        }

        if (propertyPath.localeCompare('itemSelectionProps.destinationFieldName') === 0) {
            changeCallback(propertyPath, cloneDeep((newValue as IComboBoxOption).key));
            this.context.propertyPane.refresh();
        }
    }

    protected get isRenderAsync(): boolean {
        return true;
    }

    protected async onPropertyPaneConfigurationStart() {
        await this.loadPropertyPaneResources();
    }

    /**
     * Determines the input query text value based on Dynamic Data
     */
    private _getInputQueryTextValue(): string {

        let inputQueryText: string = undefined; // {inputQueryText} token should always resolve as '' by default

        // tryGetValue() will resolve to '' if no Web Part is connected or if the connection is removed
        // The value can be also 'undefined' if the data source is not already loaded on the page.
        let inputQueryFromDataSource = "";
        if (this.wbProperties.queryText) {
            try {
                inputQueryFromDataSource = DynamicPropertyHelper.tryGetValueSafe(this.wbProperties.queryText);
                if (inputQueryFromDataSource !== undefined && typeof (inputQueryFromDataSource) === 'string') {
                    inputQueryFromDataSource = decodeURIComponent(inputQueryFromDataSource);
                }

            } catch (error) {
                // Likely issue when q=%25 in spfx
            }
        }

        if (!inputQueryFromDataSource) { // '' or 'undefined'

            if (this.wbProperties.useDefaultQueryText) {
                inputQueryText = this.wbProperties.defaultQueryText;
            } else if (inputQueryFromDataSource !== undefined) {
                inputQueryText = inputQueryFromDataSource;
            }

        } else if (typeof (inputQueryFromDataSource) === 'string') {
            inputQueryText = decodeURIComponent(inputQueryFromDataSource);
        }

        return inputQueryText;
    }

    /**
     * Loads the Microsoft Graph Toolkit library dynamically
     */
    private async loadMsGraphToolkit() {

        // Load Microsoft Graph Toolkit dynamically
        const { Providers, SharePointProvider } = await import(
            /* webpackChunkName: 'microsoft-graph-toolkit' */
            '@microsoft/mgt/dist/es6'
        );

        if (!Providers.globalProvider) {
            Providers.globalProvider = new SharePointProvider(this.context);
        }
    }

    /**
     * Loads extensions from registered extensibility librairies
     */
    private async loadExtensions(librariesConfiguration: IExtensibilityConfiguration[]) {

        // Load extensibility library if present
        const extensibilityLibraries = await this.extensibilityService.loadExtensibilityLibraries(librariesConfiguration);

        // Load extensibility additions
        if (extensibilityLibraries.length > 0) {

            // Load customizations from extensibility libraries
            extensibilityLibraries.forEach(extensibilityLibrary => {

                // Add custom layouts if any
                if (extensibilityLibrary.getCustomLayouts)
                    this.availableLayoutDefinitions = this.availableLayoutDefinitions.concat(extensibilityLibrary.getCustomLayouts());

                // Add custom web components if any
                if (extensibilityLibrary.getCustomWebComponents)
                    this.availableWebComponentDefinitions = this.availableWebComponentDefinitions.concat(extensibilityLibrary.getCustomWebComponents());

                // Registers Handlebars customizations in the local namespace
                if (extensibilityLibrary.registerHandlebarsCustomizations)
                    extensibilityLibrary.registerHandlebarsCustomizations(this.templateService.Handlebars);

            });
        }
    }

    public async loadPropertyPaneResources(): Promise<void> {

        const { PropertyFieldCodeEditor, PropertyFieldCodeEditorLanguages } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/propertyFields/codeEditor'
        );

        this._propertyFieldCodeEditor = PropertyFieldCodeEditor;
        this._propertyFieldCodeEditorLanguages = PropertyFieldCodeEditorLanguages;


        const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
        );
        this._propertyFieldCollectionData = PropertyFieldCollectionData;
        this._customCollectionFieldType = CustomCollectionFieldType;

        // Code editor component for property pane controls
        this._textDialogComponent = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '../../controls/TextDialog'
        );

        const { PropertyFieldToggleWithCallout } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/PropertyFieldToggleWithCallout'
        );

        const { PropertyPaneWebPartInformation } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation'
        );

        this._propertyPaneWebPartInformation = PropertyPaneWebPartInformation;

        const { CalloutTriggers } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/common/callout/Callout'
        );

        const { PropertyFieldNumber } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/PropertyFieldNumber'
        );

        const { PropertyPanePropertyEditor } = await import(
            /* webpackChunkName: 'pnp-modern-search-property-pane' */
            '@pnp/spfx-property-controls/lib/PropertyPanePropertyEditor'
        );
        this._propertyPanePropertyEditor = PropertyPanePropertyEditor;

        this._propertyFieldToogleWithCallout = PropertyFieldToggleWithCallout;
        this._propertyFieldCalloutTriggers = CalloutTriggers;

        this._propertyFieldNumber = PropertyFieldNumber;

        this.propertyPaneConnectionsGroup = await this.getConnectionOptionsGroup();
    }

    /**
     * Binds event fired from pagination web components
     */
    private bindPagingEvents() {

        this.domElement.addEventListener('pageNumberUpdated', ((ev: CustomEvent) => {

            // We ensure the event if not propagated outside the component (i.e. other Web Part instances)
            ev.stopImmediatePropagation();

            const eventDetails: IPageEventInfo = ev.detail;

            // These information comes from the PaginationWebComponent class
            this.currentPageNumber = eventDetails.pageNumber;
            this.currentPageLinkUrl = eventDetails.pageLink;
            this.availablePageLinks = eventDetails.pageLinks;

            this.render();

        }).bind(this));
    }

    /**
     * Initializes required Web Part properties
     */
    private initializeProperties() {
        this.wbProperties.selectedLayoutKey = this.wbProperties.selectedLayoutKey ? this.wbProperties.selectedLayoutKey : BuiltinLayoutsKeys.Cards;
        this.wbProperties.resultTypes = this.wbProperties.resultTypes ? this.wbProperties.resultTypes : [];
        this.wbProperties.dataSourceProperties = this.wbProperties.dataSourceProperties ? this.wbProperties.dataSourceProperties : {};

        if (!this.wbProperties.queryText) {
            this.wbProperties.queryText = new DynamicProperty<string>(this.context.dynamicDataProvider);
            this.wbProperties.queryText.setValue('');
        }

        this.wbProperties.queryTextSource = this.wbProperties.queryTextSource ? this.wbProperties.queryTextSource : QueryTextSource.StaticValue;
        this.wbProperties.layoutProperties = this.wbProperties.layoutProperties ? this.wbProperties.layoutProperties : {};

        // Common options 
        this.wbProperties.showSelectedFilters = this.wbProperties.showSelectedFilters !== undefined ? this.wbProperties.showSelectedFilters : false;
        this.wbProperties.showResultsCount = this.wbProperties.showResultsCount !== undefined ? this.wbProperties.showResultsCount : true;
        this.wbProperties.showBlankIfNoResult = this.wbProperties.showBlankIfNoResult !== undefined ? this.wbProperties.showBlankIfNoResult : false;
        this.wbProperties.useMicrosoftGraphToolkit = this.wbProperties.useMicrosoftGraphToolkit !== undefined ? this.wbProperties.useMicrosoftGraphToolkit : false;
        this.wbProperties.enableTelemetry = this.wbProperties.enableTelemetry !== undefined ? this.wbProperties.enableTelemetry : true;

        // Item selection properties
        if (!this.wbProperties.selectedItemFieldValue) {
            this.wbProperties.selectedItemFieldValue = new DynamicProperty<string>(this.context.dynamicDataProvider);
            this.wbProperties.selectedItemFieldValue.setValue('');
        }

        this.wbProperties.itemSelectionProps = this.wbProperties.itemSelectionProps !== undefined ? this.wbProperties.itemSelectionProps : {
            allowItemSelection: false,
            destinationFieldName: undefined,
            selectionMode: ItemSelectionMode.AsDataFilter,
            allowMulti: false,
            valuesOperator: FilterConditionOperator.OR
        };

        this.wbProperties.extensibilityLibraryConfiguration = this.wbProperties.extensibilityLibraryConfiguration ? this.wbProperties.extensibilityLibraryConfiguration : [{
            name: commonStrings.General.Extensibility.DefaultExtensibilityLibraryName,
            enabled: true,
            id: Constants.DEFAULT_EXTENSIBILITY_LIBRARY_COMPONENT_ID
        }];

        if (this.wbProperties.selectedVerticalKeys === undefined) {
            this.wbProperties.selectedVerticalKeys = [];
        }

        // Adapt to new schema since 4.1.0
        if (this.wbProperties['selectedVerticalKey'] && this.wbProperties.selectedVerticalKeys.indexOf(this.wbProperties['selectedVerticalKey']) === -1) {
            this.wbProperties.selectedVerticalKeys.push(this.wbProperties['selectedVerticalKey']);
        }

        this.wbProperties.useVerticals = this.wbProperties.useVerticals !== undefined ? this.wbProperties.useVerticals : false;

        if (!this.wbProperties.paging) {

            this.wbProperties.paging = {
                itemsCountPerPage: 10,
                pagingRange: 5,
                showPaging: true,
                hideDisabled: true,
                hideFirstLastPages: false,
                hideNavigation: false,
                useNextLinks: false
            };
        }
    }

    /**
     * Returns property pane 'Styling' page groups
     */
    private getStylingPageGroups(): IPropertyPaneGroup[] {

        const canEditTemplate = this.wbProperties.externalTemplateUrl && this.wbProperties.selectedLayoutKey === BuiltinLayoutsKeys.ResultsCustom ? false : true;

        let stylingFields: IPropertyPaneField<any>[] = [
            PropertyPaneChoiceGroup('selectedLayoutKey', {
                options: LayoutHelper.getLayoutOptions(this.availableLayoutDefinitions)
            })
        ];

        let resultTypeInlineTemplate = undefined;

        switch (this.wbProperties.selectedLayoutKey) {
            case BuiltinLayoutsKeys.SimpleList:
                resultTypeInlineTemplate = require('../../layouts/resultTypes/default_simple_list.html');
                break;

            case BuiltinLayoutsKeys.Cards:
                resultTypeInlineTemplate = require('../../layouts/resultTypes/default_cards.html');
                break;

            case BuiltinLayoutsKeys.ResultsCustom:
                resultTypeInlineTemplate = require('../../layouts/resultTypes/default_custom.html');
                break;

            case BuiltinLayoutsKeys.People:
                resultTypeInlineTemplate = require('../../layouts/resultTypes/default_people.html');
                break;

            default:
                break;
        }

        // We can customize the template for any layout
        stylingFields.push(
            this._propertyFieldCodeEditor('inlineTemplateContent', {
                label: webPartStrings.PropertyPane.LayoutPage.DialogButtonLabel,
                panelTitle: webPartStrings.PropertyPane.LayoutPage.DialogTitle,
                initialValue: this.templateContentToDisplay,
                deferredValidationTime: 500,
                onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                properties: this.wbProperties,
                disabled: !canEditTemplate,
                key: 'inlineTemplateContentCodeEditor',
                language: this._propertyFieldCodeEditorLanguages.Handlebars
            }),
            this._propertyFieldCollectionData('resultTypes', {
                manageBtnLabel: webPartStrings.PropertyPane.LayoutPage.ResultTypes.EditResultTypesLabel,
                key: 'resultTypes',
                panelHeader: webPartStrings.PropertyPane.LayoutPage.ResultTypes.EditResultTypesLabel,
                panelDescription: webPartStrings.PropertyPane.LayoutPage.ResultTypes.ResultTypesDescription,
                enableSorting: true,
                label: webPartStrings.PropertyPane.LayoutPage.ResultTypes.ResultTypeslabel,
                value: this.wbProperties.resultTypes,
                disabled: this.wbProperties.selectedLayoutKey === BuiltinLayoutsKeys.DetailsList
                    || this.wbProperties.selectedLayoutKey === BuiltinLayoutsKeys.ResultsDebug
                    || this.wbProperties.selectedLayoutKey === BuiltinLayoutsKeys.Slider ? true : false,
                fields: [
                    {
                        id: 'property',
                        title: webPartStrings.PropertyPane.LayoutPage.ResultTypes.ConditionPropertyLabel,
                        type: this._customCollectionFieldType.dropdown,
                        required: true,
                        options: this._currentDataResultsSourceData.availableFieldsFromResults.map(field => {
                            return {
                                key: field,
                                text: field
                            };
                        })
                    },
                    {
                        id: 'operator',
                        title: webPartStrings.PropertyPane.LayoutPage.ResultTypes.CondtionOperatorValue,
                        type: this._customCollectionFieldType.dropdown,
                        defaultValue: ResultTypeOperator.Equal,
                        required: true,
                        options: [
                            {
                                key: ResultTypeOperator.Equal,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.EqualOperator
                            },
                            {
                                key: ResultTypeOperator.NotEqual,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.NotEqualOperator
                            },
                            {
                                key: ResultTypeOperator.Contains,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.ContainsOperator
                            },
                            {
                                key: ResultTypeOperator.StartsWith,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.StartsWithOperator
                            },
                            {
                                key: ResultTypeOperator.NotNull,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.NotNullOperator
                            },
                            {
                                key: ResultTypeOperator.GreaterOrEqual,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.GreaterOrEqualOperator
                            },
                            {
                                key: ResultTypeOperator.GreaterThan,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.GreaterThanOperator
                            },
                            {
                                key: ResultTypeOperator.LessOrEqual,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.LessOrEqualOperator
                            },
                            {
                                key: ResultTypeOperator.LessThan,
                                text: webPartStrings.PropertyPane.LayoutPage.ResultTypes.LessThanOperator
                            }
                        ]
                    },
                    {
                        id: 'value',
                        title: webPartStrings.PropertyPane.LayoutPage.ResultTypes.ConditionValueLabel,
                        type: this._customCollectionFieldType.string,
                        required: false,
                    },
                    {
                        id: "inlineTemplateContent",
                        title: webPartStrings.PropertyPane.LayoutPage.ResultTypes.InlineTemplateContentLabel,
                        type: this._customCollectionFieldType.custom,
                        onCustomRender: ((field, value, onUpdate) => {
                            return (
                                React.createElement("div", null,
                                    React.createElement(this._textDialogComponent.TextDialog, {
                                        language: this._propertyFieldCodeEditorLanguages.Handlebars,
                                        dialogTextFieldValue: value ? value : resultTypeInlineTemplate,
                                        onChanged: (fieldValue) => onUpdate(field.id, fieldValue),
                                        strings: {
                                            cancelButtonText: webPartStrings.PropertyPane.LayoutPage.ResultTypes.CancelButtonText,
                                            dialogButtonText: webPartStrings.PropertyPane.LayoutPage.ResultTypes.DialogButtonText,
                                            dialogTitle: webPartStrings.PropertyPane.LayoutPage.ResultTypes.DialogTitle,
                                            saveButtonText: webPartStrings.PropertyPane.LayoutPage.ResultTypes.SaveButtonText
                                        }
                                    })
                                )
                            );
                        }).bind(this)
                    },
                    {
                        id: 'externalTemplateUrl',
                        title: webPartStrings.PropertyPane.LayoutPage.ResultTypes.ExternalUrlLabel,
                        type: this._customCollectionFieldType.url,
                        onGetErrorMessage: this.onTemplateUrlChange.bind(this),
                        placeholder: 'https://mysite/Documents/external.html'
                    }
                ]
            })
        );

        // Only show the template external URL for 'Custom' option
        if (this.wbProperties.selectedLayoutKey === BuiltinLayoutsKeys.ResultsCustom) {
            stylingFields.push(
                PropertyPaneTextField('externalTemplateUrl', {
                    label: webPartStrings.PropertyPane.LayoutPage.TemplateUrlFieldLabel,
                    placeholder: webPartStrings.PropertyPane.LayoutPage.TemplateUrlPlaceholder,
                    deferredValidationTime: 500,
                    validateOnFocusIn: true,
                    validateOnFocusOut: true,
                    onGetErrorMessage: this.onTemplateUrlChange.bind(this)
                }));
        }

        let groups: IPropertyPaneGroup[] = [
            {
                groupName: webPartStrings.PropertyPane.LayoutPage.LayoutSelectionGroupName,
                groupFields: stylingFields
            }
        ];

        let layoutOptionsFields: IPropertyPaneField<any>[] = [
            PropertyPaneToggle('itemSelectionProps.allowItemSelection', {
                label: webPartStrings.PropertyPane.LayoutPage.AllowItemSelection
            }),
            PropertyPaneToggle('showBlankIfNoResult', {
                label: webPartStrings.PropertyPane.LayoutPage.ShowBlankIfNoResult,
            }),
            PropertyPaneToggle('showResultsCount', {
                label: webPartStrings.PropertyPane.LayoutPage.ShowResultsCount,
            }),
            PropertyPaneToggle('useMicrosoftGraphToolkit', {
                label: webPartStrings.PropertyPane.LayoutPage.UseMicrosoftGraphToolkit,
            })
        ];

        if (this.wbProperties.filtersDataSourceReference) {
            layoutOptionsFields.push(
                PropertyPaneToggle('showSelectedFilters', {
                    label: webPartStrings.PropertyPane.LayoutPage.ShowSelectedFilters,
                })
            );
        }

        if (this.wbProperties.itemSelectionProps.allowItemSelection) {

            layoutOptionsFields.splice(1, 0,
                PropertyPaneToggle('itemSelectionProps.allowMulti', {
                    label: webPartStrings.PropertyPane.LayoutPage.AllowMultipleItemSelection,
                }),
                PropertyPaneHorizontalRule()
            );
        }

        // Add template options if any
        const layoutOptions = this.getLayoutTemplateOptions();

        groups.push(
            {
                groupName: webPartStrings.PropertyPane.LayoutPage.CommonOptionsGroupName,
                groupFields: layoutOptionsFields
            },
            {
                groupName: webPartStrings.PropertyPane.LayoutPage.LayoutTemplateOptionsGroupName,
                groupFields: layoutOptions
            }
        );

        return groups;
    }

    /**
     * Returns property pane 'Paging' group fields
     */
    private getPagingGroupFields(): IPropertyPaneField<any>[] {

        let groupFields: IPropertyPaneField<any>[] = [];

        if (this.dataSource) {

            // Only show paging option if the data source supports it (dynamic or static)
            if (this.dataSource.getPagingBehavior()) {

                groupFields.push(
                    PropertyPaneToggle('paging.showPaging', {
                        label: webPartStrings.PropertyPane.DataSourcePage.ShowPagingFieldName,
                    }),
                    this._propertyFieldNumber('paging.itemsCountPerPage', {
                        label: webPartStrings.PropertyPane.DataSourcePage.ItemsCountPerPageFieldName,
                        maxValue: 500,
                        minValue: 1,
                        value: this.wbProperties.paging.itemsCountPerPage,
                        disabled: !this.wbProperties.paging.showPaging,
                        key: 'paging.itemsCountPerPage'
                    }),
                    PropertyPaneSlider('paging.pagingRange', {
                        label: webPartStrings.PropertyPane.DataSourcePage.PagingRangeFieldName,
                        max: 50,
                        min: 0, // 0 = no page numbers displayed
                        step: 1,
                        showValue: true,
                        value: this.wbProperties.paging.pagingRange,
                        disabled: !this.wbProperties.paging.showPaging
                    }),
                    PropertyPaneHorizontalRule(),
                    PropertyPaneToggle('paging.hideNavigation', {
                        label: webPartStrings.PropertyPane.DataSourcePage.HideNavigationFieldName,
                        disabled: !this.wbProperties.paging.showPaging
                    }),
                    PropertyPaneToggle('paging.hideFirstLastPages', {
                        label: webPartStrings.PropertyPane.DataSourcePage.HideFirstLastPagesFieldName,
                        disabled: !this.wbProperties.paging.showPaging
                    }),
                    PropertyPaneToggle('paging.hideDisabled', {
                        label: webPartStrings.PropertyPane.DataSourcePage.HideDisabledFieldName,
                        disabled: !this.wbProperties.paging.showPaging
                    })
                );
            }
        }

        return groupFields;
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
                                        }).bind(this),
                                        styles: { text: { position: "absolute", top: 0, left: "38px", right: 0, bottom: 0, display: "flex", alignItems: "center", width: "fit-content", cursor: "pointer" } }
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
     * Builds the data source options list from the available data sources
     */
    private getDataSourceOptions(): IPropertyPaneChoiceGroupOption[] {

        let dataSourceOptions: IPropertyPaneChoiceGroupOption[] = [];

        this.availableDataSourceDefinitions.forEach((source) => {
            dataSourceOptions.push({
                iconProps: {
                    officeFabricIconFontName: source.iconName
                },
                imageSize: {
                    width: 200,
                    height: 100
                },
                key: source.key,
                text: source.name,
            });
        });

        return dataSourceOptions;
    }

    /**
     * Returns layout template options if any
     */
    private getLayoutTemplateOptions(): IPropertyPaneField<any>[] {

        if (this.layout && !this.errorMessage) {
            return this.layout.getPropertyPaneFieldsConfiguration(this._currentDataResultsSourceData.availableFieldsFromResults);
        } else {
            return [];
        }
    }

    private getTemplateSlotOptions(): IPropertyPaneField<any>[] {

        let templateSlotFields: IPropertyPaneField<any>[] = [];
        if (this.dataSource) {

            let availableOptions: IComboBoxOption[];
            if (this._currentDataResultsSourceData.availableFieldsFromResults.length > 0) {
                availableOptions = this._currentDataResultsSourceData.availableFieldsFromResults.map(field => {
                    return {
                        key: field,
                        text: field
                    };
                });
            }
            else {
                availableOptions = this.dataSource.getTemplateSlots().map(slot => {
                    return {
                        key: slot.slotField,
                        text: slot.slotField
                    };
                });
            }

            templateSlotFields.push(
                this._propertyFieldCollectionData('templateSlots', {
                    manageBtnLabel: webPartStrings.PropertyPane.DataSourcePage.TemplateSlots.ConfigureSlotsBtnLabel,
                    key: 'templateSlots',
                    enableSorting: false,
                    panelHeader: webPartStrings.PropertyPane.DataSourcePage.TemplateSlots.ConfigureSlotsPanelHeader,
                    panelDescription: webPartStrings.PropertyPane.DataSourcePage.TemplateSlots.ConfigureSlotsPanelDescription,
                    label: webPartStrings.PropertyPane.DataSourcePage.TemplateSlots.ConfigureSlotsLabel,
                    value: this.wbProperties.templateSlots,
                    fields: [
                        {
                            id: 'slotName',
                            title: webPartStrings.PropertyPane.DataSourcePage.TemplateSlots.SlotNameFieldName,
                            type: this._customCollectionFieldType.string
                        },
                        {
                            id: 'slotField',
                            title: webPartStrings.PropertyPane.DataSourcePage.TemplateSlots.SlotFieldFieldName,
                            type: this._customCollectionFieldType.custom,
                            required: false,
                            onCustomRender: (field, value, onUpdate, item) => {
                                return (
                                    React.createElement("div", null,
                                        React.createElement(AsyncCombo, {
                                            allowFreeform: true,
                                            availableOptions: availableOptions,
                                            placeholder: webPartStrings.PropertyPane.DataSourcePage.TemplateSlots.SlotFieldPlaceholderName,
                                            textDisplayValue: item[field.id] ? item[field.id] : '',
                                            defaultSelectedKey: item[field.id] ? item[field.id] : '',
                                            onLoadOptions: () => {
                                                return Promise.resolve(availableOptions);
                                            },
                                            onUpdateOptions: () => { },
                                            onUpdate: (filterValue: IComboBoxOption) => {
                                                onUpdate(field.id, filterValue.key);
                                            }
                                        } as IAsyncComboProps)
                                    )
                                );
                            }
                        }
                    ]
                })
            );
        }

        return templateSlotFields;
    }

    private getSearchQueryTextFields(): IPropertyPaneField<any>[] {
        let searchQueryTextFields: IPropertyPaneField<any>[] = [
            this._propertyFieldToogleWithCallout('useInputQueryText', {
                label: webPartStrings.PropertyPane.ConnectionsPage.UseInputQueryText,
                calloutTrigger: this._propertyFieldCalloutTriggers.Hover,
                key: 'useInputQueryText',
                calloutContent: React.createElement('p', { style: { maxWidth: 250, wordBreak: 'break-word' } }, webPartStrings.PropertyPane.ConnectionsPage.UseInputQueryTextHoverMessage),
                onText: commonStrings.General.OnTextLabel,
                offText: commonStrings.General.OffTextLabel,
                checked: this.wbProperties.useInputQueryText
            })
        ];

        if (this.wbProperties.useInputQueryText) {

            searchQueryTextFields.push(
                PropertyPaneChoiceGroup('queryTextSource', {
                    options: [
                        {
                            key: QueryTextSource.StaticValue,
                            text: webPartStrings.PropertyPane.ConnectionsPage.InputQueryTextStaticValue
                        },
                        {
                            key: QueryTextSource.DynamicValue,
                            text: webPartStrings.PropertyPane.ConnectionsPage.InputQueryTextDynamicValue
                        }
                    ]
                })
            );

            switch (this.wbProperties.queryTextSource) {

                case QueryTextSource.StaticValue:
                    searchQueryTextFields.push(
                        PropertyPaneTextField('queryText', {
                            label: webPartStrings.PropertyPane.ConnectionsPage.SearchQueryTextFieldLabel,
                            description: webPartStrings.PropertyPane.ConnectionsPage.SearchQueryTextFieldDescription,
                            multiline: true,
                            resizable: true,
                            placeholder: webPartStrings.PropertyPane.ConnectionsPage.SearchQueryPlaceHolderText,
                            onGetErrorMessage: this._validateEmptyField.bind(this),
                            deferredValidationTime: 500
                        })
                    );
                    break;

                case QueryTextSource.DynamicValue:
                    searchQueryTextFields.push(
                        PropertyPaneDynamicField('queryText', {
                            label: ''
                        }),
                        PropertyPaneCheckbox('useDefaultQueryText', {
                            text: webPartStrings.PropertyPane.ConnectionsPage.SearchQueryTextUseDefaultQuery,
                            disabled: this.wbProperties.queryText.reference === undefined
                        })
                    );

                    if (this.wbProperties.useDefaultQueryText && this.wbProperties.queryText.reference !== undefined) {
                        searchQueryTextFields.push(
                            PropertyPaneTextField('defaultQueryText', {
                                label: webPartStrings.PropertyPane.ConnectionsPage.SearchQueryTextDefaultValue,
                                multiline: true
                            })
                        );
                    }

                    break;

                default:
                    break;
            }
        }

        return searchQueryTextFields;
    }

    private async getDataResultsConnectionFields(): Promise<IPropertyPaneField<any>[]> {

        let dataResultsConnectionFields: IPropertyPaneField<any>[] = [
            PropertyPaneToggle('useDynamicFiltering', {
                label: webPartStrings.PropertyPane.ConnectionsPage.UseDynamicFilteringsWebPartLabel,
                checked: this.wbProperties.useDynamicFiltering
            })
        ];

        if (this.wbProperties.useDynamicFiltering) {

            let isSourceFieldConfigured: boolean = false;

            // Make sure a property is selected in the source according to the reference format.
            // Ex: PageContext:UrlData:queryParameters.q = Page environment
            // Ex: WebPart.544c1372-42df-47c3-94d6-017428cd2baf.1272b161-3435-4815-99a1-996590334cff:AvailableFieldValuesFromResults:FileType = Search Results
            if (this.wbProperties.selectedItemFieldValue.reference) {
                isSourceFieldConfigured = /^.+:.+:(.+)$/.test(this.wbProperties.selectedItemFieldValue.reference);
            }

            dataResultsConnectionFields.push(

                // Allow both 'Search Results' Web Parts and OOTB SharePoint List Web Parts 
                PropertyPaneDynamicFieldSet({
                    label: webPartStrings.PropertyPane.ConnectionsPage.UseDataResultsFromComponentsLabel,
                    fields: [
                        PropertyPaneDynamicField('selectedItemFieldValue', {
                            label: webPartStrings.PropertyPane.ConnectionsPage.UseDataResultsFromComponentsLabel,
                        })
                    ],
                    sharedConfiguration: {
                        depth: DynamicDataSharedDepth.Property,
                        property: {
                            filters: {
                                propertyId: DynamicDataProperties.AvailableFieldValuesFromResults
                            }
                        }
                    }
                })
            );

            if (isSourceFieldConfigured) {

                const availableOptions: IComboBoxOption[] = this._currentDataResultsSourceData.availableFieldsFromResults.map(field => {
                    return {
                        key: field,
                        text: field
                    };
                });

                dataResultsConnectionFields.splice(4, 0,
                    new PropertyPaneAsyncCombo('itemSelectionProps.destinationFieldName', {
                        label: webPartStrings.PropertyPane.ConnectionsPage.SourceDestinationFieldLabel,
                        availableOptions: availableOptions,
                        description: webPartStrings.PropertyPane.ConnectionsPage.SourceDestinationFieldDescription,
                        allowMultiSelect: false,
                        allowFreeform: true,
                        searchAsYouType: false,
                        defaultSelectedKeys: this.wbProperties.selectedVerticalKeys,
                        textDisplayValue: this.wbProperties.itemSelectionProps.destinationFieldName,
                        onPropertyChange: this.onCustomPropertyUpdate.bind(this),
                    })
                );
            }

            if (isSourceFieldConfigured && this.wbProperties.itemSelectionProps.destinationFieldName) {

                dataResultsConnectionFields.splice(4, 0,
                    PropertyPaneChoiceGroup('itemSelectionProps.selectionMode', {
                        options: [
                            {
                                key: ItemSelectionMode.AsDataFilter,
                                text: webPartStrings.PropertyPane.LayoutPage.AsDataFiltersSelectionMode
                            },
                            {
                                key: ItemSelectionMode.AsTokenValue,
                                text: webPartStrings.PropertyPane.LayoutPage.AsTokensSelectionMode
                            }
                        ],
                        label: webPartStrings.PropertyPane.LayoutPage.SelectionModeLabel,
                    })
                );

                if (this.wbProperties.itemSelectionProps.selectionMode === ItemSelectionMode.AsDataFilter) {
                    dataResultsConnectionFields.splice(5, 0,
                        this._propertyPaneWebPartInformation({
                            description: `<em>${webPartStrings.PropertyPane.LayoutPage.AsDataFiltersDescription}</em>`,
                            key: 'selectionModeText'
                        }),
                        PropertyPaneChoiceGroup('itemSelectionProps.valuesOperator', {
                            options: [
                                {
                                    key: FilterConditionOperator.OR,
                                    text: 'OR'
                                },
                                {
                                    key: FilterConditionOperator.AND,
                                    text: 'AND'
                                },
                            ],
                            label: webPartStrings.PropertyPane.LayoutPage.FilterValuesOperator
                        })
                    );
                } else {
                    dataResultsConnectionFields.splice(4, 0,
                        this._propertyPaneWebPartInformation({
                            description: `<em>${webPartStrings.PropertyPane.LayoutPage.AsTokensDescription}</em>`,
                            key: 'selectionModeText'
                        })
                    );
                }
            }
        }

        return dataResultsConnectionFields;
    }

    private async getFiltersConnectionFields(): Promise<IPropertyPaneField<any>[]> {

        let filtersConnectionFields: IPropertyPaneField<any>[] = [
            PropertyPaneToggle('useFilters', {
                label: webPartStrings.PropertyPane.ConnectionsPage.UseFiltersWebPartLabel,
                checked: this.wbProperties.useFilters
            })
        ];

        if (this.wbProperties.useFilters) {
            filtersConnectionFields.splice(1, 0,
                PropertyPaneDropdown('filtersDataSourceReference', {
                    options: await this.dynamicDataService.getAvailableDataSourcesByType(ComponentType.SearchFilters),
                    label: webPartStrings.PropertyPane.ConnectionsPage.UseFiltersFromComponentLabel
                })
            );
        }

        return filtersConnectionFields;
    }

    private async getVerticalsConnectionFields(): Promise<IPropertyPaneField<any>[]> {

        let verticalsConnectionFields: IPropertyPaneField<any>[] = [
            PropertyPaneToggle('useVerticals', {
                label: webPartStrings.PropertyPane.ConnectionsPage.UseSearchVerticalsWebPartLabel,
                checked: this.wbProperties.useVerticals
            })
        ];

        if (this.wbProperties.useVerticals) {
            verticalsConnectionFields.splice(1, 0,
                PropertyPaneDropdown('verticalsDataSourceReference', {
                    options: await this.dynamicDataService.getAvailableDataSourcesByType(ComponentType.SearchVerticals),
                    label: webPartStrings.PropertyPane.ConnectionsPage.UseSearchVerticalsFromComponentLabel
                })
            );

            if (this.wbProperties.verticalsDataSourceReference) {

                // Get all available verticals
                if (this._verticalsConnectionSourceData) {
                    const availableVerticals = DynamicPropertyHelper.tryGetValueSafe(this._verticalsConnectionSourceData);

                    if (availableVerticals) {

                        // Get the corresponding text for selected keys
                        let selectedKeysAsText: string[] = [];

                        availableVerticals.verticalsConfiguration.forEach(verticalConfiguration => {
                            if (this.wbProperties.selectedVerticalKeys.indexOf(verticalConfiguration.key) !== -1) {
                                selectedKeysAsText.push(verticalConfiguration.tabName);
                            }
                        });

                        verticalsConnectionFields.push(
                            new PropertyPaneAsyncCombo('selectedVerticalKeys', {
                                availableOptions: availableVerticals.verticalsConfiguration.filter(v => !v.isLink).map(verticalConfiguration => {
                                    return {
                                        key: verticalConfiguration.key,
                                        text: verticalConfiguration.tabName
                                    };
                                }),
                                allowMultiSelect: true,
                                allowFreeform: false,
                                description: webPartStrings.PropertyPane.ConnectionsPage.LinkToVerticalLabelHoverMessage,
                                label: webPartStrings.PropertyPane.ConnectionsPage.LinkToVerticalLabel,
                                searchAsYouType: false,
                                defaultSelectedKeys: this.wbProperties.selectedVerticalKeys,
                                textDisplayValue: selectedKeysAsText.join(','),
                                onPropertyChange: this.onCustomPropertyUpdate.bind(this),
                            }),
                        );
                    }
                }
            }
        }

        return verticalsConnectionFields;
    }

    private async getConnectionOptionsGroup(): Promise<IPropertyPaneGroup[]> {

        const filterConnectionFields = await this.getFiltersConnectionFields();
        const verticalConnectionFields = await this.getVerticalsConnectionFields();
        const dataResultsConnectionsFields = await this.getDataResultsConnectionFields();

        let availableConnectionsGroup: IPropertyPaneGroup[] = [
            {
                groupName: webPartStrings.PropertyPane.ConnectionsPage.ConnectionsPageGroupName,
                groupFields: [
                    ...this.getSearchQueryTextFields(),
                    PropertyPaneHorizontalRule(),
                    ...filterConnectionFields,
                    PropertyPaneHorizontalRule(),
                    ...verticalConnectionFields,
                    PropertyPaneHorizontalRule(),
                    ...dataResultsConnectionsFields
                ]
            }
        ];

        return availableConnectionsGroup;
    }

    /**
     * Gets the data source instance according to the current selected one
     * @param dataSourceKey the selected data source provider key
     * @param dataSourceDefinitions the available source definitions
     * @returns the data source provider instance
     */
    private async getDataSourceInstance(dataSourceKey: string): Promise<IDataSource> {

        let dataSource: IDataSource = undefined;
        let serviceKey: ServiceKey<IDataSource> = undefined;

        if (dataSourceKey) {

            // If it is a builtin data source, we load the corresponding known class file asynchronously for performance purpose
            // We also create the service key at the same time to be able to get an instance
            switch (dataSourceKey) {

                // SharePoint Search API
                case BuiltinDataSourceProviderKeys.SharePointSearch:

                    const { SharePointSearchDataSource } = await import(
                        /* webpackChunkName: 'pnp-modern-search-sharepoint-search-datasource' */
                        '../../dataSources/SharePointSearchDataSource'
                    );

                    serviceKey = ServiceKey.create<IDataSource>('ModernSearch:SharePointSearchDataSource', SharePointSearchDataSource);
                    break;

                // Microsoft Search API
                case BuiltinDataSourceProviderKeys.MicrosoftSearch:

                    const { MicrosoftSearchDataSource } = await import(
                        /* webpackChunkName: 'pnp-modern-search-microsoft-search-datasource' */
                        '../../dataSources/MicrosoftSearchDataSource'
                    );

                    serviceKey = ServiceKey.create<IDataSource>('ModernSearch:SharePointSearchDataSource', MicrosoftSearchDataSource);
                    break;

                // Azure Function Search API
                case BuiltinDataSourceProviderKeys.AzureSearch:

                    const { AzureSearchDataSource } = await import(
                        /* webpackChunkName: 'pnp-modern-search-azure-search-datasource' */
                        '../../dataSources/AzureSearchDataSource'
                    );

                    serviceKey = ServiceKey.create<IDataSource>('ModernSearch:SharePointSearchDataSource', AzureSearchDataSource);
                    break;

                default:
                    break;
            }

            return new Promise<IDataSource>((resolve, reject) => {

                // Register here services we want to expose to custom data sources (ex: TokenService)
                // The instances are shared across all data sources. It means when properties will be set once for all consumers. Be careful manipulating these instance properties. 
                const childServiceScope = ServiceScopeHelper.registerChildServices(this.webPartInstanceServiceScope, [
                    serviceKey,
                    TaxonomyService.ServiceKey,
                    SharePointSearchService.ServiceKey,
                    TokenService.ServiceKey
                ]);

                childServiceScope.whenFinished(async () => {

                    this.tokenService = childServiceScope.consume<ITokenService>(TokenService.ServiceKey);

                    // Initialize the token values
                    await this.setTokens();

                    // Register the data source service in the Web Part scope only (child scope of the current scope)
                    dataSource = childServiceScope.consume<IDataSource>(serviceKey);

                    // Verifiy if the data source implements correctly the IDataSource interface and BaseDataSource methods
                    const isValidDataSource = (dataSourceInstance: IDataSource): dataSourceInstance is BaseDataSource<any> => {
                        return (
                            (dataSourceInstance as BaseDataSource<any>).getAppliedFilters !== undefined &&
                            (dataSourceInstance as BaseDataSource<any>).getData !== undefined &&
                            (dataSourceInstance as BaseDataSource<any>).getFilterBehavior !== undefined &&
                            (dataSourceInstance as BaseDataSource<any>).getItemCount !== undefined &&
                            (dataSourceInstance as BaseDataSource<any>).getPagingBehavior !== undefined &&
                            (dataSourceInstance as BaseDataSource<any>).getPropertyPaneGroupsConfiguration !== undefined &&
                            (dataSourceInstance as BaseDataSource<any>).getTemplateSlots !== undefined &&
                            (dataSourceInstance as BaseDataSource<any>).onInit !== undefined &&
                            (dataSourceInstance as BaseDataSource<any>).onPropertyUpdate !== undefined
                        );
                    };

                    if (!isValidDataSource(dataSource)) {
                        reject(new Error(Text.format(commonStrings.General.Extensibility.InvalidDataSourceInstance, dataSourceKey)));
                    }

                    // Initialize the data source with current Web Part properties
                    if (dataSource) {
                        // Initializes Web part lifecycle methods and properties
                        dataSource.properties = this.wbProperties.dataSourceProperties;
                        dataSource.context = this.context;
                        dataSource.editMode = this.displayMode == DisplayMode.Edit;
                        dataSource.render = this.render;

                        // Initializes available services
                        dataSource.serviceKeys = {
                            TokenService: TokenService.ServiceKey
                        };

                        await dataSource.onInit();

                        // Initialize slots
                        if (isEmpty(this.wbProperties.templateSlots)) {
                            this.wbProperties.templateSlots = dataSource.getTemplateSlots();
                            this._defaultTemplateSlots = dataSource.getTemplateSlots();
                        }

                        resolve(dataSource);
                    }
                });
            });
        }
    }

    /**
      * Custom handler when the external template file URL
      * @param value the template file URL value
      */
    private async onTemplateUrlChange(value: string): Promise<string> {

        try {
            // Doesn't raise any error if file is empty (otherwise error message will show on initial load...)
            if (isEmpty(value)) {
                return Promise.resolve('');
            }
            // Resolves an error if the file isn't a valid .htm or .html file
            else if (!this.templateService.isValidTemplateFile(value)) {
                return Promise.resolve(webPartStrings.PropertyPane.LayoutPage.ErrorTemplateExtension);
            }
            // Resolves an error if the file doesn't answer a simple head request
            else {
                await this.templateService.ensureFileResolves(value);
                return Promise.resolve('');
            }
        } catch (error) {
            return Promise.resolve(Text.format(webPartStrings.PropertyPane.LayoutPage.ErrorTemplateResolve, error));
        }
    }

    /**
     * Initializes the template according to the property pane current configuration
     * @returns the template content as a string
     */
    private async initTemplate(): Promise<void> {

        // Gets the template content according to the selected key
        const selectedLayoutTemplateContent = this.availableLayoutDefinitions.filter(layout => { return layout.key === this.wbProperties.selectedLayoutKey; })[0].templateContent;

        if (this.wbProperties.selectedLayoutKey === BuiltinLayoutsKeys.ResultsCustom) {

            if (this.wbProperties.externalTemplateUrl) {
                this.templateContentToDisplay = await this.templateService.getFileContent(this.wbProperties.externalTemplateUrl);
            } else {
                this.templateContentToDisplay = this.wbProperties.inlineTemplateContent ? this.wbProperties.inlineTemplateContent : selectedLayoutTemplateContent;
            }

        } else {
            this.templateContentToDisplay = selectedLayoutTemplateContent;
        }

        // Register result types inside the template      
        await this.templateService.registerResultTypes(this.wbProperties.resultTypes);

        return;
    }

    /**
      * Initializes the service scope manager singleton instance
      * The scopes whithin the solution are as follow 
      *   Top root scope (Shared by all client side components)     
      *   |--- ExtensibilityService
      *   |--- DateHelper
      *   |--- Client side component scope (i.e shared with all Web Part instances)
      *     |--- SPHttpClient
      *     |--- <other SPFx http services>
      *     |--- (Web Part Scope (created with startNewChild))
      *       |--- DynamicDataService
      *       |--- TemplateService
      *       |--- (Data Source scope)
      *         |--- SharePointSearchDataSource
      *         |--- TokenService, SearchService, etc.       
    */
    private initializeWebPartServices(): void {

        // Register specific Web Part service instances
        this.webPartInstanceServiceScope = this.context.serviceScope.startNewChild();
        this.templateService = this.webPartInstanceServiceScope.createAndProvide(TemplateService.ServiceKey, TemplateService);
        this.dynamicDataService = this.webPartInstanceServiceScope.createAndProvide(DynamicDataService.ServiceKey, DynamicDataService);
        this.dynamicDataService.dynamicDataProvider = this.context.dynamicDataProvider;
        this.webPartInstanceServiceScope.finish();
    }

    /**
     * Set token values from Web Part property bag
     */
    private async setTokens() {

        if (this.tokenService) {

            // Input query text
            const inputQueryText = this._getInputQueryTextValue();
            this.tokenService.setTokenValue(BuiltinTokenNames.inputQueryText, inputQueryText);

            // Legacy token for SharePoint and Microsoft Search data sources
            this.tokenService.setTokenValue(BuiltinTokenNames.searchTerms, inputQueryText);

            // Selected filters
            if (this._filtersConnectionSourceData) {

                const filtersSourceData: IDataFilterSourceData = DynamicPropertyHelper.tryGetValueSafe(this._filtersConnectionSourceData);

                if (filtersSourceData) {

                    // Set the token as 'null' if no filter is selected meaning the token is available but with no data (different from 'undefined')
                    // It is the caller responsability to check if the value is empty before using it in an expression (ex: `if(empty('{filters}'),'doA','doB)`)
                    let filterTokens: IDataFilterToken = null;

                    const allValues = flatten(filtersSourceData.selectedFilters.map(f => f.values));

                    // Make sure we have values in selected filters
                    if (filtersSourceData.selectedFilters.length > 0 && !isEmpty(allValues)) {

                        filterTokens = {};

                        // Build the initial structure for the configured filter names
                        filtersSourceData.filterConfiguration.forEach(filterConfiguration => {

                            // Initialize to an empty object so the token service can resolve it to an empty string instead leaving the token '{filters}' as is
                            filterTokens[filterConfiguration.filterName] = null;
                        });

                        filtersSourceData.selectedFilters.forEach(filter => {

                            const configuration = DataFilterHelper.getConfigurationForFilter(filter, filtersSourceData.filterConfiguration);

                            if (configuration) {

                                let filterTokenValue: IDataFilterTokenValue = null;

                                const filterValues = filter.values.map(value => value.value).join(',');

                                // Don't tokenize the filter if there is no value.
                                if (filterValues.length > 0) {
                                    filterTokenValue = {
                                        valueAsText: filterValues
                                    };
                                }

                                if (configuration.selectedTemplate === BuiltinFilterTemplates.DateRange) {

                                    let fromDate = undefined;
                                    let toDate = undefined;

                                    // Determine start and end dates by operator
                                    filter.values.forEach(filterValue => {
                                        if (filterValue.operator === FilterComparisonOperator.Geq && !fromDate) {
                                            fromDate = filterValue.value;
                                        }

                                        if (filterValue.operator === FilterComparisonOperator.Lt && !toDate) {
                                            toDate = fromDate = filterValue.value;
                                        }
                                    });

                                    filterTokenValue.fromDate = fromDate;
                                    filterTokenValue.toDate = toDate;
                                }

                                filterTokens[filter.filterName] = filterTokenValue;
                            }
                        });
                    }

                    this.tokenService.setTokenValue(BuiltinTokenNames.filters, filterTokens);
                }
            }

            // Current selected Search Results or SharePoint List Web Part
            const destinationFieldName = this.wbProperties.itemSelectionProps.destinationFieldName;

            const itemFieldValues: string[] = DynamicPropertyHelper.tryGetValuesSafe(this.wbProperties.selectedItemFieldValue);

            if (destinationFieldName) {

                let filterTokens = {
                    [destinationFieldName]: {
                        valueAsText: null,
                    } as IDataFilterTokenValue
                };

                filterTokens[destinationFieldName].valueAsText = itemFieldValues.length > 0 ? itemFieldValues.join(',') : undefined;  // This allow the `{? <KQL expression>}` to work
                this.tokenService.setTokenValue(BuiltinTokenNames.filters, filterTokens);
            }

            // Current selected vertical
            if (this._verticalsConnectionSourceData) {
                const verticalSourceData = DynamicPropertyHelper.tryGetValueSafe(this._verticalsConnectionSourceData);

                // Tokens for verticals are resolved first locally in the Search Verticals WP itself. If some tokens are not recognized in the string (ex: undefined in their TokenService instance), they will be left untounched. 
                // In this case, we need to resolve them in the current Search Results WP context as they only exist here (ex: itemsCountPerPage)
                if (verticalSourceData && verticalSourceData.selectedVertical) {
                    const resolvedSelectedVertical: IDataVertical = {
                        key: verticalSourceData.selectedVertical.key,
                        name: verticalSourceData.selectedVertical.name,
                        value: await this.tokenService.resolveTokens(verticalSourceData.selectedVertical.value)
                    };

                    this.tokenService.setTokenValue(BuiltinTokenNames.verticals, resolvedSelectedVertical);
                }
            }
        }
    }

    /**
     * Make sure the dynamic properties are correctly connected to the corresponding sources according to the proeprty pane settings
     */
    private ensureDynamicDataSourcesConnection() {

        // Filters Web Part data source
        if (this.wbProperties.filtersDataSourceReference) {

            if (!this._filtersConnectionSourceData) {
                this._filtersConnectionSourceData = new DynamicProperty<IDataFilterSourceData>(this.context.dynamicDataProvider);
            }

            this._filtersConnectionSourceData.setReference(this.wbProperties.filtersDataSourceReference);
            this._filtersConnectionSourceData.register(this.render);

        } else {

            if (this._filtersConnectionSourceData) {
                this._filtersConnectionSourceData.unregister(this.render);
            }
        }

        // Verticals Web Part data source
        if (this.wbProperties.verticalsDataSourceReference) {

            if (!this._verticalsConnectionSourceData) {
                this._verticalsConnectionSourceData = new DynamicProperty<IDataVerticalSourceData>(this.context.dynamicDataProvider);
            }

            this._verticalsConnectionSourceData.setReference(this.wbProperties.verticalsDataSourceReference);
            this._verticalsConnectionSourceData.register(this.render);

        } else {
            if (this._verticalsConnectionSourceData) {
                this._verticalsConnectionSourceData.unregister(this.render);
            }
        }

    }

    /**
     * Checks if a field if empty or not
     * @param value the value to check
     */
    private _validateEmptyField(value: string): string {

        if (!value) {
            return commonStrings.General.EmptyFieldErrorMessage;
        }

        return '';
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
     * Reset the paging information for PagingBehavior.Dynamic data sources
     */
    private _resetPagingData() {
        this.availablePageLinks = [];
        this.currentPageLinkUrl = null;
    }

    /**
   * Get the data context to be passed to the data source according to current connections/configurations
   */
    private getDataContext(): IDataContext {

        // Input query text
        const inputQueryText = this._getInputQueryTextValue();

        // Build the data context to pass to the data source
        let dataContext: IDataContext = {
            pageNumber: this.currentPageNumber,
            itemsCountPerPage: this.wbProperties.paging.itemsCountPerPage,
            paging: {
                nextLinkUrl: this.currentPageLinkUrl,
                pageLinks: this.availablePageLinks
            },
            filters: {
                selectedFilters: [],
                filtersConfiguration: [],
                instanceId: undefined,
                filterOperator: undefined
            },
            verticals: {
                selectedVertical: undefined
            },
            inputQueryText: inputQueryText,
            queryStringParameters: UrlHelper.getQueryStringParams()
        };

        // Connected Search Results or SharePoint List Web Part
        const itemFieldValues: string[] = DynamicPropertyHelper.tryGetValuesSafe(this.wbProperties.selectedItemFieldValue);

        if (itemFieldValues && itemFieldValues.length > 0 && this.wbProperties.itemSelectionProps.destinationFieldName) {

            // Set the selected items to the data context. This will force data to be fetched again
            dataContext.selectedItemValues = itemFieldValues;

            // Convert the current selection into search filters format, just like the Data Filter Web Part
            if (this.wbProperties.itemSelectionProps.selectionMode === ItemSelectionMode.AsDataFilter) {

                const filterValues: IDataFilterValue[] = uniq(itemFieldValues) // Remove duplicate values selected by the user
                    .filter(value => !value || typeof value === 'string')
                    .map(fieldValue => {
                        return {
                            name: fieldValue,
                            value: fieldValue,
                            operator: FilterComparisonOperator.Eq
                        };
                    });
                if (filterValues.length > 0) {
                    dataContext.filters.selectedFilters.push({
                        filterName: this.wbProperties.itemSelectionProps.destinationFieldName,
                        values: filterValues,
                        operator: this.wbProperties.itemSelectionProps.valuesOperator
                    });
                }
            }
        }

        // Connected Search Filters
        if (this._filtersConnectionSourceData) {
            const filtersSourceData: IDataFilterSourceData = DynamicPropertyHelper.tryGetValueSafe(this._filtersConnectionSourceData);
            if (filtersSourceData) {

                // Reset the page number if filters have been updated by the user
                if (!isEqual(filtersSourceData.selectedFilters, this._lastSelectedFilters)) {
                    dataContext.pageNumber = 1;
                    this.currentPageNumber = 1;
                    this._resetPagingData();
                }

                // Use the filter confiugration and then get the corresponding values 
                dataContext.filters.filtersConfiguration = filtersSourceData.filterConfiguration;
                dataContext.filters.selectedFilters = dataContext.filters.selectedFilters.concat(filtersSourceData.selectedFilters);
                dataContext.filters.filterOperator = filtersSourceData.filterOperator;
                dataContext.filters.instanceId = filtersSourceData.instanceId;

                this._lastSelectedFilters = dataContext.filters.selectedFilters;
            }
        }

        // Connected Search Verticals
        if (this._verticalsConnectionSourceData) {
            const verticalsSourceData: IDataVerticalSourceData = DynamicPropertyHelper.tryGetValueSafe(this._verticalsConnectionSourceData);
            if (verticalsSourceData) {
                dataContext.verticals.selectedVertical = verticalsSourceData.selectedVertical;
            }
        }

        // If input query text changes, then we need to reset the paging
        if (!isEqual(inputQueryText, this._lastInputQueryText)) {
            dataContext.pageNumber = 1;
            this.currentPageNumber = 1;
            this._resetPagingData();
        }

        this._lastInputQueryText = inputQueryText;

        return dataContext;
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

    /**
     * Handler when data are retreived from the source
     * @param availableFields the available fields
     * @param filters the available filters from the data source
     * @param pageNumber the current page number
     * @param nextLinkUrl the next link URL if any
     * @param pageLinks the page links
     */
    private _onDataRetrieved(availableDataSourceFields: string[], filters?: IDataFilterResult[], pageNumber?: number, nextLinkUrl?: string, pageLinks?: string[]) {

        this._currentDataResultsSourceData.availableFieldsFromResults = availableDataSourceFields;
        this.currentPageNumber = pageNumber;
        this.availablePageLinks = pageLinks;
        this.currentPageLinkUrl = nextLinkUrl;

        // Set the available filters from the data source 
        if (filters) {
            this._currentDataResultsSourceData.availablefilters = filters;
        }

        // Check if the Web part is connected to a data vertical
        if (this._verticalsConnectionSourceData && this.wbProperties.selectedVerticalKeys.length > 0) {
            const verticalData = DynamicPropertyHelper.tryGetValueSafe(this._verticalsConnectionSourceData);

            // For edit mode only, we want to see the data
            if (verticalData && this.wbProperties.selectedVerticalKeys.indexOf(verticalData.selectedVertical.key) === -1 && this.displayMode === DisplayMode.Read) {

                // If the current selected vertical is not the one configured for this Web Part, we reset
                // the data soure information since we don't want to expose them to consumers
                this._currentDataResultsSourceData = {
                    availableFieldsFromResults: [],
                    availablefilters: []
                };
            }
        }

        // Notfify dynamic data consumers data have changed
        if (this.context && this.context.dynamicDataSourceManager && !this.context.dynamicDataSourceManager.isDisposed) {
            this.context.dynamicDataSourceManager.notifyPropertyChanged(ComponentType.SearchResults);
        }

        // Extra call to refresh the property pane in the case where data sources rely on results fields in there configuration (ex: ODataDataSource)
        if (this.context && this.context.propertyPane) {
            this.context.propertyPane.refresh();
        }
    }

    /**
     * Handler when an item is selected in the results 
     * @param currentSelectedItems the current selected items
     */
    private _onItemSelected(currentSelectedItems: { [key: string]: any }[]) {

        this._currentDataResultsSourceData.selectedItems = cloneDeep(currentSelectedItems);

        // Notfify dynamic data consumers data have changed
        this.context.dynamicDataSourceManager.notifyPropertyChanged(DynamicDataProperties.AvailableFieldValuesFromResults);
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
}