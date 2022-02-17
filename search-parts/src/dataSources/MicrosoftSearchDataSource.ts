import { BaseDataSource, FilterSortType, FilterSortDirection, ITemplateSlot, BuiltinTemplateSlots, IDataContext, ITokenService, FilterBehavior, PagingBehavior, IDataFilterResult, IDataFilterResultValue, FilterComparisonOperator } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneLabel, IPropertyPaneField, PropertyPaneToggle, PropertyPaneHorizontalRule, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { cloneDeep, isEmpty } from '@microsoft/sp-lodash-subset';
import { TokenService } from "../services/tokenService/TokenService";
import { Guid, ServiceScope } from '@microsoft/sp-core-library';
import { IComboBoxOption } from 'office-ui-fabric-react';
import { PropertyPaneAsyncCombo } from "../controls/PropertyPaneAsyncCombo/PropertyPaneAsyncCombo";
import * as commonStrings from 'CommonStrings';
import { IMicrosoftSearchRequest, ISearchRequestAggregation, SearchAggregationSortBy, ISearchSortProperty, IMicrosoftSearchQuery, IQueryAlterationOptions, ICustomAadApplicationOptions } from '../models/search/IMicrosoftSearchRequest';
import { DateHelper } from '../helpers/DateHelper';
import { DataFilterHelper } from "../helpers/DataFilterHelper";
import { ISortFieldConfiguration, SortFieldDirection } from '../models/search/ISortFieldConfiguration';
import { AsyncCombo } from "../controls/PropertyPaneAsyncCombo/components/AsyncCombo";
import { IAsyncComboProps } from "../controls/PropertyPaneAsyncCombo/components/IAsyncComboProps";
import { PropertyPaneNonReactiveTextField } from "../controls/PropertyPaneNonReactiveTextField/PropertyPaneNonReactiveTextField";
import { ISharePointSearchService } from "../services/searchService/ISharePointSearchService";
import { SharePointSearchService } from "../services/searchService/SharePointSearchService";
import { IMicrosoftSearchDataSourceData } from "../models/search/IMicrosoftSearchDataSourceData";

import * as React from "react";
import { BuiltinDataSourceProviderKeys } from "./AvailableDataSources";
import { IMicrosoftSearchService } from "../services/searchService/IMicrosoftSearchService";
import { MicrosoftSearchService } from "../services/searchService/MicrosoftSearchService";
import { UrlHelper } from "../helpers/UrlHelper";
import { ISite } from "../models/common/ISIte";
import { ITaxonomyService } from "../services/taxonomyService/ITaxonomyService";
import { TaxonomyService } from "../services/taxonomyService/TaxonomyService";
import { ITermInfo } from "@pnp/sp/taxonomy";

export enum EntityType {
    Message = 'message',
    Event = 'event',
    Drive = 'drive',
    DriveItem = 'driveItem',
    ExternalItem = 'externalItem',
    List = 'list',
    ListItem = 'listItem',
    Site = 'site',
    Person = 'person'
}

export interface IMicrosoftSearchDataSourceProperties {

    /**
     * The entity types to search. See for the complete list
     */
    entityTypes: EntityType[];


    /**
     * Contains the fields to be returned for each resource object specified in entityTypes, allowing customization of the fields returned by default otherwise, including additional fields such as custom managed properties from SharePoint and OneDrive, or custom fields in externalItem from content ingested by Graph connectors.
     */
    fields: string[];

    /**
     * The sort fields configuration
     */
    sortProperties: ISortFieldConfiguration[];

    /**
     * This triggers hybrid sort for messages : the first 3 messages are the most relevant. This property is only applicable to entityType=message
     */
    enableTopResults: boolean;

    /**
     * The content sources for external items
     */
    contentSourceConnectionIds: string[];

    /**
     * The query alteration options for spelling corrections
     */
    queryAlterationOptions: IQueryAlterationOptions;

    /**
    * The search query template
    */
    queryTemplate: string;

    /**
    * Flag indicating if the Microsoft Search beta endpoint should be used
     */
    useBetaEndpoint: boolean;

    useCustomAadApplication: boolean;

    customAadApplicationOptions: ICustomAadApplicationOptions;
}

export class MicrosoftSearchDataSource extends BaseDataSource<IMicrosoftSearchDataSourceProperties> {

    private _tokenService: ITokenService;
    private _sharePointSearchService: ISharePointSearchService;
    private _microsoftSearchService: IMicrosoftSearchService;
    private _taxonomyService: ITaxonomyService;

    private _propertyPaneWebPartInformation: any = null;
    private _availableFields: IComboBoxOption[] = [];
    private _microsoftSearchUrl: string;
    private _availableManagedProperties: IComboBoxOption[] = [];

    private _availableEntityTypeOptions: IComboBoxOption[] = [
        {
            key: EntityType.Message,
            text: "Messages"
        },
        {
            key: EntityType.Event,
            text: "Events"
        },
        {
            key: EntityType.Drive,
            text: "Drive"
        },
        {
            key: EntityType.DriveItem,
            text: "Drive Items"
        },
        {
            key: EntityType.ExternalItem,
            text: "External Items"
        },
        {
            key: EntityType.ListItem,
            text: "List Items"
        },
        {
            key: EntityType.List,
            text: "List"
        },
        {
            key: EntityType.Site,
            text: "Sites"
        },
        {
            key: EntityType.Person,
            text: "People"
        }
    ];

    /**
     * The data source items count
     */
    private _itemsCount: number = 0;

    /**
     * A date helper instance
     */
    private dateHelper: DateHelper;

    /**
    * The moment.js library reference
    */
    private moment: any;

    private _propertyFieldCollectionData: any = null;
    private _customCollectionFieldType: any = null;

    public constructor(serviceScope: ServiceScope) {
        super(serviceScope);

        serviceScope.whenFinished(() => {
            this._tokenService = serviceScope.consume<ITokenService>(TokenService.ServiceKey);
            this._sharePointSearchService = serviceScope.consume<ISharePointSearchService>(SharePointSearchService.ServiceKey);
            this._microsoftSearchService = serviceScope.consume<IMicrosoftSearchService>(MicrosoftSearchService.ServiceKey);
            this._taxonomyService = serviceScope.consume<ITaxonomyService>(TaxonomyService.ServiceKey);
        });
    }

    public async onInit(): Promise<void> {

        this.dateHelper = this.serviceScope.consume<DateHelper>(DateHelper.ServiceKey);
        this.moment = await this.dateHelper.moment();

        if (this.editMode) {
            // Use the same chunk name as the main Web Part to avoid recreating/loading a new one
            const { PropertyFieldCollectionData, CustomCollectionFieldType } = await import(
                /* webpackChunkName: 'pnp-modern-search-property-pane' */
                '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'
            );

            const { PropertyPaneWebPartInformation } = await import(
                /* webpackChunkName: 'pnp-modern-search-property-pane' */
                '@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation'
            );

            this._propertyPaneWebPartInformation = PropertyPaneWebPartInformation;
            this._propertyFieldCollectionData = PropertyFieldCollectionData;
            this._customCollectionFieldType = CustomCollectionFieldType;
        }

        await this.initProperties();
    }

    public getItemCount(): number {
        return this._itemsCount;
    }

    public getFilterBehavior(): FilterBehavior {
        return FilterBehavior.Dynamic;
    }

    public getPagingBehavior(): PagingBehavior {
        return PagingBehavior.Dynamic;
    }

    public async getData(dataContext: IDataContext): Promise<IMicrosoftSearchDataSourceData> {

        let results: IMicrosoftSearchDataSourceData = {
            items: []
        };

        // Ensuring at least one entity type is selected before launching a search
        if (this._properties.entityTypes.length > 0) {
            const searchQuery = await this.buildMicrosoftSearchQuery(dataContext);
            results = await this.search(searchQuery);
        } else {
            // If no entity is selected, manually set the results to prevent
            // having the previous search results items count displayed.
            this._itemsCount = 0;
        }

        return results;
    }

    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

        const entityTypesDisplayValue = this._availableEntityTypeOptions.map((option) => {
            if (this.properties.entityTypes.indexOf(option.key as EntityType) !== -1) {
                return option.text;
            }
        });
        let selectFieldsFields: IPropertyPaneField<any>[] = [];
        let contentSourceConnectionIdsFields: IPropertyPaneField<any>[] = [];
        let enableTopResultsFields: IPropertyPaneField<any>[] = [];
        let sortPropertiesFields: IPropertyPaneField<any>[] = [];
        let queryAlterationFields: IPropertyPaneField<any>[] = [];
        let commonFields: IPropertyPaneField<any>[] = [
            PropertyPaneLabel('', {
                text: commonStrings.DataSources.MicrosoftSearch.QueryTextFieldLabel
            }),
            this._propertyPaneWebPartInformation({
                description: `<em>${commonStrings.DataSources.MicrosoftSearch.QueryTextFieldInfoMessage}</em>`,
                key: 'queryText'
            }),
            new PropertyPaneNonReactiveTextField('dataSourceProperties.queryTemplate', {
                componentKey: `${BuiltinDataSourceProviderKeys.MicrosoftSearch}-queryTemplate`,
                defaultValue: this.properties.queryTemplate,
                label: commonStrings.DataSources.MicrosoftSearch.QueryTemplateFieldLabel,
                placeholderText: commonStrings.DataSources.MicrosoftSearch.QueryTemplatePlaceHolderText,
                multiline: true,
                description: commonStrings.DataSources.MicrosoftSearch.QueryTemplateFieldDescription,
                applyBtnText: commonStrings.DataSources.MicrosoftSearch.ApplyQueryTemplateBtnText,
                allowEmptyValue: false,
                rows: 8
            }),
            new PropertyPaneAsyncCombo('dataSourceProperties.entityTypes', {
                availableOptions: this._availableEntityTypeOptions,
                allowMultiSelect: true,
                allowFreeform: false,
                description: "",
                label: commonStrings.DataSources.MicrosoftSearch.EntityTypesField,
                placeholder: "",
                searchAsYouType: false,
                defaultSelectedKeys: this.properties.entityTypes,
                onPropertyChange: this.onCustomPropertyUpdate.bind(this),
                textDisplayValue: entityTypesDisplayValue.filter(e => e).join(",")
            }),
            new PropertyPaneAsyncCombo('dataSourceProperties.fields', {
                availableOptions: this._availableManagedProperties,
                allowMultiSelect: true,
                allowFreeform: true,
                description: commonStrings.DataSources.MicrosoftSearch.SelectedFieldsPropertiesFieldDescription,
                label: commonStrings.DataSources.MicrosoftSearch.SelectedFieldsPropertiesFieldLabel,
                placeholder: commonStrings.DataSources.MicrosoftSearch.SelectedFieldsPlaceholderLabel,
                searchAsYouType: false,
                defaultSelectedKeys: this.properties.fields,
                onLoadOptions: this.getAvailableProperties.bind(this),
                onPropertyChange: this.onCustomPropertyUpdate.bind(this),
                onUpdateOptions: ((options: IComboBoxOption[]) => {
                    this._availableFields = options;
                }).bind(this)
            })
        ];

        let customAadPropertiesFields: IPropertyPaneField<any>[] = [
            PropertyPaneToggle('dataSourceProperties.useCustomAadApplication', {
                label: commonStrings.DataSources.MicrosoftSearch.UseCustomAadApplication
            })
        ];

        if (this.properties.useCustomAadApplication) {
            customAadPropertiesFields.push(
                PropertyPaneTextField('dataSourceProperties.customAadApplicationOptions.tenantId', {
                    label: commonStrings.DataSources.MicrosoftSearch.TenantIdFieldLabel,
                    placeholder: commonStrings.DataSources.MicrosoftSearch.TenantIdPlaceholder,
                    description: commonStrings.DataSources.MicrosoftSearch.TenantIdFieldDescription
                }),
                PropertyPaneTextField('dataSourceProperties.customAadApplicationOptions.clientId', {
                    label: commonStrings.DataSources.MicrosoftSearch.ClientIdFieldLabel,
                    placeholder: commonStrings.DataSources.MicrosoftSearch.ClientIdPlaceholder
                }),
                PropertyPaneTextField('dataSourceProperties.customAadApplicationOptions.redirectUrl', {
                    label: commonStrings.DataSources.MicrosoftSearch.RedirectUrlFieldLabel,
                    placeholder: commonStrings.DataSources.MicrosoftSearch.RedirectUrlPlaceholder
                })
            );
        }

        let useBetaEndpointFields: IPropertyPaneField<any>[] = [
            PropertyPaneHorizontalRule(),
            PropertyPaneToggle('dataSourceProperties.useBetaEndpoint', {
                label: commonStrings.DataSources.MicrosoftSearch.UseBetaEndpoint
            })
        ];

        // Sorting results is currently only supported on the following SharePoint and OneDrive types: driveItem, listItem, list, site.
        if (this.properties.entityTypes.indexOf(EntityType.DriveItem) !== -1 ||
            this.properties.entityTypes.indexOf(EntityType.ListItem) !== -1 ||
            this.properties.entityTypes.indexOf(EntityType.Site) !== -1 ||
            this.properties.entityTypes.indexOf(EntityType.List) !== -1) {

            sortPropertiesFields.push(
                this._propertyFieldCollectionData('dataSourceProperties.sortProperties', {
                    manageBtnLabel: commonStrings.DataSources.SearchCommon.Sort.EditSortLabel,
                    key: 'sortProperties',
                    enableSorting: true,
                    panelHeader: commonStrings.DataSources.SearchCommon.Sort.EditSortLabel,
                    panelDescription: commonStrings.DataSources.SearchCommon.Sort.SortListDescription,
                    label: commonStrings.DataSources.SearchCommon.Sort.SortPropertyPaneFieldLabel,
                    value: this.properties.sortProperties,
                    fields: [
                        {
                            id: 'sortField',
                            title: commonStrings.DataSources.SearchCommon.Sort.SortFieldColumnLabel,
                            type: this._customCollectionFieldType.custom,
                            required: true,
                            onCustomRender: ((field, value, onUpdate, item, itemId, onError) => {

                                return React.createElement("div", { key: `${field.id}-${itemId}` },
                                    React.createElement(AsyncCombo, {
                                        defaultSelectedKey: item[field.id] ? item[field.id] : '',
                                        onUpdate: (option: IComboBoxOption) => {

                                            this._sharePointSearchService.validateSortableProperty(option.key as string).then((sortable: boolean) => {
                                                if (!sortable) {
                                                    onError(field.id, commonStrings.DataSources.SearchCommon.Sort.SortInvalidSortableFieldMessage);
                                                } else {
                                                    onUpdate(field.id, option.key as string);
                                                    onError(field.id, '');
                                                }
                                            });
                                        },
                                        allowMultiSelect: false,
                                        allowFreeform: true,
                                        availableOptions: this._availableManagedProperties,
                                        onLoadOptions: this.getAvailableProperties.bind(this),
                                        onUpdateOptions: ((options: IComboBoxOption[]) => {
                                            this._availableManagedProperties = options;
                                        }).bind(this),
                                        placeholder: commonStrings.DataSources.SearchCommon.Sort.SortFieldColumnPlaceholder,
                                        useComboBoxAsMenuWidth: false // Used when screen resolution is too small to display the complete value  
                                    } as IAsyncComboProps));
                            }).bind(this)
                        },
                        {
                            id: 'sortDirection',
                            title: commonStrings.DataSources.SearchCommon.Sort.SortDirectionColumnLabel,
                            type: this._customCollectionFieldType.dropdown,
                            required: false,
                            options: [
                                {
                                    key: SortFieldDirection.Ascending,
                                    text: commonStrings.DataSources.SearchCommon.Sort.SortDirectionAscendingLabel
                                },
                                {
                                    key: SortFieldDirection.Descending,
                                    text: commonStrings.DataSources.SearchCommon.Sort.SortDirectionDescendingLabel
                                }
                            ],
                            defaultValue: SortFieldDirection.Ascending
                        }
                    ]
                })
            );
        }

        if (this.properties.entityTypes.indexOf(EntityType.ExternalItem) !== -1) {
            contentSourceConnectionIdsFields.push(
                new PropertyPaneAsyncCombo('dataSourceProperties.contentSourceConnectionIds', {
                    availableOptions: [],
                    allowMultiSelect: true,
                    allowFreeform: true,
                    description: commonStrings.DataSources.MicrosoftSearch.ContentSourcesFieldDescriptionLabel,
                    label: commonStrings.DataSources.MicrosoftSearch.ContentSourcesFieldLabel,
                    placeholder: commonStrings.DataSources.MicrosoftSearch.ContentSourcesFieldPlaceholderLabel,
                    searchAsYouType: false,
                    defaultSelectedKeys: this.properties.contentSourceConnectionIds,
                    onPropertyChange: this.onCustomPropertyUpdate.bind(this)
                })
            );
        }

        if (this.properties.entityTypes.indexOf(EntityType.Message) !== -1 && this.properties.entityTypes.length === 1) {
            enableTopResultsFields.push(PropertyPaneToggle('dataSourceProperties.enableTopResults', {
                label: commonStrings.DataSources.MicrosoftSearch.EnableTopResultsLabel
            }));
        }

        // https://docs.microsoft.com/en-us/graph/search-concept-speller#known-limitations
        if (this.properties.useBetaEndpoint &&
            (this.properties.entityTypes.indexOf(EntityType.Message) !== -1 ||
                this.properties.entityTypes.indexOf(EntityType.Event) !== -1 ||
                this.properties.entityTypes.indexOf(EntityType.Site) !== -1 ||
                this.properties.entityTypes.indexOf(EntityType.Drive) !== -1 ||
                this.properties.entityTypes.indexOf(EntityType.DriveItem) !== -1 ||
                this.properties.entityTypes.indexOf(EntityType.List) !== -1 ||
                this.properties.entityTypes.indexOf(EntityType.ListItem) !== -1 ||
                this.properties.entityTypes.indexOf(EntityType.ExternalItem) !== -1)) {
            queryAlterationFields.push(
                PropertyPaneToggle('dataSourceProperties.queryAlterationOptions.enableSuggestion', {
                    label: commonStrings.DataSources.MicrosoftSearch.EnableSuggestionLabel,
                    checked: this.properties.queryAlterationOptions.enableSuggestion
                }),
                PropertyPaneToggle('dataSourceProperties.queryAlterationOptions.enableModification', {
                    label: commonStrings.DataSources.MicrosoftSearch.EnableModificationLabel,
                    checked: this.properties.queryAlterationOptions.enableModification
                })
            );
        }

        let groupFields: IPropertyPaneField<any>[] = [
            ...commonFields,
            ...selectFieldsFields,
            ...sortPropertiesFields,
            ...enableTopResultsFields,
            ...contentSourceConnectionIdsFields,
            ...customAadPropertiesFields,
            ...useBetaEndpointFields,
            ...queryAlterationFields
        ];

        return [
            {
                groupName: commonStrings.DataSources.MicrosoftSearch.SourceConfigurationGroupName,
                groupFields: groupFields
            }
        ];
    }

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any) {

        if (propertyPath.localeCompare('dataSourceProperties.useBetaEndpoint') === 0) {

            if (newValue) {
                this._microsoftSearchUrl = "https://graph.microsoft.com/beta/search/query";

                // Reset beta options
                this.properties.queryAlterationOptions.enableSuggestion = false;
                this.properties.queryAlterationOptions.enableModification = false;

            } else {
                this._microsoftSearchUrl = "https://graph.microsoft.com/v1.0/search/query";
            }
        }

        if (propertyPath.localeCompare('dataSourceProperties.useCustomAadApplication') === 0) {

            if (newValue) {
                // Reset custom aad options
                this.properties.customAadApplicationOptions.tenantId = undefined;
                this.properties.customAadApplicationOptions.clientId = undefined;
                this.properties.customAadApplicationOptions.redirectUrl = undefined;
            }
        }
    }

    public onCustomPropertyUpdate(propertyPath: string, newValue: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {

        if (propertyPath.localeCompare('dataSourceProperties.entityTypes') === 0) {
            this.properties.entityTypes = (cloneDeep(newValue) as IComboBoxOption[]).map(v => { return v.key as EntityType; });
            this.context.propertyPane.refresh();
            this.render();
        }

        if (propertyPath.localeCompare('dataSourceProperties.fields') === 0) {
            let options = this.parseAndCleanOptions((cloneDeep(newValue) as IComboBoxOption[]));
            this.properties.fields = options.map(v => { return v.key as string; });
            this.context.propertyPane.refresh();
            this.render();
        }

        if (propertyPath.localeCompare('dataSourceProperties.contentSourceConnectionIds') === 0) {
            this.properties.contentSourceConnectionIds = (cloneDeep(newValue) as IComboBoxOption[]).map(v => { return v.key as string; });
            this.context.propertyPane.refresh();
            this.render();
        }
    }

    public getTemplateSlots(): ITemplateSlot[] {
        return [
            {
                slotName: BuiltinTemplateSlots.Title,
                slotField: 'resource.fields.title'
            },
            {
                slotName: BuiltinTemplateSlots.Path,
                slotField: 'resource.webUrl'
            },
            {
                slotName: BuiltinTemplateSlots.Summary,
                slotField: 'summary'
            },
            {
                slotName: BuiltinTemplateSlots.Author,
                slotField: 'resource.fields.authorOWSUSER'
            },
            {
                slotName: BuiltinTemplateSlots.FileType,
                slotField: 'resource.fields.filetype'
            },
            {
                slotName: BuiltinTemplateSlots.PreviewImageUrl,
                slotField: 'AutoPreviewImageUrl' // Field added automatically
            },
            {
                slotName: BuiltinTemplateSlots.LegacyPreviewImageUrl,
                slotField: 'resource.fields.serverRedirectedPreviewURL' // Field added automatically
            },
            {
                slotName: BuiltinTemplateSlots.PreviewUrl,
                slotField: 'AutoPreviewUrl' // Field added automatically
            },
            {
                slotName: BuiltinTemplateSlots.PreviewViewUrl,
                slotField: 'AutoPreviewViewUrl' // Field added automatically
            },
            {
                slotName: BuiltinTemplateSlots.LegacyPreviewUrl,
                slotField: 'resource.fields.serverRedirectedEmbedURL' // Field added automatically
            },
            {
                slotName: BuiltinTemplateSlots.Tags,
                slotField: 'owstaxidmetadataalltagsinfo'
            },
            {
                slotName: BuiltinTemplateSlots.Date,
                slotField: 'created'
            },
            {
                slotName: BuiltinTemplateSlots.SiteUrl,
                slotField: 'resource.fields.spSiteURL'
            },
            {
                slotName: BuiltinTemplateSlots.SiteId,
                slotField: 'resource.fields.normSiteID'
            },
            {
                slotName: BuiltinTemplateSlots.WebId,
                slotField: 'resource.fields.normWebID'
            },
            {
                slotName: BuiltinTemplateSlots.ListId,
                slotField: 'resource.fields.normListID'
            },
            {
                slotName: BuiltinTemplateSlots.ItemId,
                slotField: 'resource.fields.normUniqueID'
            },
            {
                slotName: BuiltinTemplateSlots.IsFolder,
                slotField: 'resource.fields.contentTypeId'
            }
        ];
    }

    private async initProperties(): Promise<void> {
        this.properties.entityTypes = this.properties.entityTypes !== undefined ? this.properties.entityTypes : [EntityType.DriveItem];

        const CommonFields = ["name", "title", "webUrl", "filetype", "createdBy", "createdDateTime", "lastModifiedDateTime", "parentReference", "size", "description", "file", "folder", "subject", "bodyPreview", "replyTo", "from", "sender", "start", "end", "displayName", "givenName", "surname", "userPrincipalName", "phones", "department", "ServerRedirectedPreviewURL", "ServerRedirectedEmbedURL", "owstaxIdDepartments", "normSiteID", "normWebID", "normListID", "normUniqueID", "contentTypeId"];

        this.properties.fields = this.properties.fields !== undefined ? this.properties.fields : CommonFields;
        this.properties.sortProperties = this.properties.sortProperties !== undefined ? this.properties.sortProperties : [];
        this.properties.contentSourceConnectionIds = this.properties.contentSourceConnectionIds !== undefined ? this.properties.contentSourceConnectionIds : [];

        this.properties.queryAlterationOptions = this.properties.queryAlterationOptions ?? { enableModification: false, enableSuggestion: false };
        this.properties.queryTemplate = this.properties.queryTemplate ? this.properties.queryTemplate : "{searchTerms}";
        this.properties.useBetaEndpoint = this.properties.useBetaEndpoint !== undefined ? this.properties.useBetaEndpoint : false;
        this.properties.useCustomAadApplication = this.properties.useCustomAadApplication !== undefined ? this.properties.useCustomAadApplication : false;

        const queryStringParameters: { [parameter: string]: string } = UrlHelper.getQueryStringParams();
        if (queryStringParameters && !isEmpty(queryStringParameters["scope"]) && !isEmpty(queryStringParameters["sid"])) {
            const siteId = queryStringParameters["sid"];
            const site: ISite = await this._microsoftSearchService.getSiteBySiteId(siteId);
            this.properties.queryTemplate += ` Path:${site.webUrl} `;
        }

        if (this.properties.useBetaEndpoint) {
            this._microsoftSearchUrl = "https://graph.microsoft.com/beta/search/query";
        } else {
            this._microsoftSearchUrl = "https://graph.microsoft.com/v1.0/search/query";
        }
    }

    private async buildMicrosoftSearchQuery(dataContext: IDataContext): Promise<IMicrosoftSearchQuery> {

        let searchQuery: IMicrosoftSearchQuery = {
            requests: [],
            queryAlterationOptions: {
                enableModification: this.properties.queryAlterationOptions.enableModification,
                enableSuggestion: this.properties.queryAlterationOptions.enableSuggestion
            }
        };
        let aggregations: ISearchRequestAggregation[] = [];
        let aggregationFilters: string[] = [];
        let sortProperties: ISearchSortProperty[] = [];
        let contentSources: string[] = [];
        let queryText = '*'; // Default query string if not specified, the API does not support empty value
        let from = 0;
        let queryTemplate: string = this.properties.queryTemplate;

        // Query text
        if (dataContext.inputQueryText) {
            queryText = await this._tokenService.resolveTokens(dataContext.inputQueryText);
        }

        if (dataContext.filters.selectedFilters.length > 0) {
            if (dataContext.filters.selectedFilters.filter(selectedFilter => selectedFilter.values.length > 0 && selectedFilter.filterName === "relatedHubSites").length > 0) {

                let termSetId: string = null;

                // Get filter config for hub sites filter
                const hubSitesFilterConfig = dataContext.filters.filtersConfiguration.filter(filterConfig => filterConfig.selectedTemplate === "TaxonomyPickerFilterTemplate" && filterConfig.filterName === "relatedHubSites");

                if (hubSitesFilterConfig && hubSitesFilterConfig.length === 1) {
                    termSetId = hubSitesFilterConfig[0].termSetId;
                }

                if (!isEmpty(termSetId)) {
                    const hubSiteFilter = dataContext.filters.selectedFilters[0];

                    var promises: Promise<ITermInfo>[] = hubSiteFilter.values.map(async (filterValue) => {
                        let termId = filterValue.value;
                        const termInfo = await this._taxonomyService.getTermById(Guid.parse(termSetId), Guid.parse(termId));
                        return new Promise<ITermInfo>((resolve, reject) => resolve(termInfo));
                    });

                    var results: Promise<ITermInfo[]> = Promise.all(promises);
                    const terms: ITermInfo[] = await results as ITermInfo[];

                    const hubSiteQueryTemplates = terms.map((term) => {
                        if (term && term.properties && term.properties.length > 0) {
                            const hubSiteIdProperty = term.properties.filter(o => o.key === "SiteId");
                            if (hubSiteIdProperty && hubSiteIdProperty.length === 1) {
                                const hubSiteId = hubSiteIdProperty[0].value;
                                return `(DepartmentId:{${hubSiteId}} OR DepartmentId:${hubSiteId} OR RelatedHubSites:${hubSiteId})`;
                            }
                        }
                    }).filter(c => c);

                    if (!isEmpty(queryTemplate.trim()) && hubSiteQueryTemplates && hubSiteQueryTemplates.length > 0) {
                        queryTemplate += hubSiteQueryTemplates.join(' OR ');
                    }
                }
            }
        }

        // Query modification
        queryTemplate = await this._tokenService.resolveTokens(queryTemplate);
        if (!isEmpty(queryTemplate.trim())) {

            // Use {searchTerms} or {inputQueryText} to use orginal value
            queryText = queryTemplate.trim();
        }

        // Paging
        if (dataContext.pageNumber > 1) {
            from = (dataContext.pageNumber - 1) * dataContext.itemsCountPerPage;
        }

        // Build aggregations
        aggregations = dataContext.filters.filtersConfiguration.map(filterConfig => {

            let aggregation: ISearchRequestAggregation = {
                field: filterConfig.filterName,
                bucketDefinition: {
                    isDescending: filterConfig.sortDirection === FilterSortDirection.Ascending ? false : true,
                    minimumCount: 0,
                    sortBy: filterConfig.sortBy === FilterSortType.ByCount ? SearchAggregationSortBy.Count : SearchAggregationSortBy.KeyAsString
                },
                size: filterConfig && filterConfig.maxBuckets ? filterConfig.maxBuckets : 10
            };

            if (filterConfig.selectedTemplate === "DateIntervalFilterTemplate") {

                const pastYear = this.moment(new Date()).subtract(1, 'years').subtract('minutes', 1).toISOString();
                const past3Months = this.moment(new Date()).subtract(3, 'months').subtract('minutes', 1).toISOString();
                const pastMonth = this.moment(new Date()).subtract(1, 'months').subtract('minutes', 1).toISOString();
                const pastWeek = this.moment(new Date()).subtract(1, 'week').subtract('minutes', 1).toISOString();
                const past24hours = this.moment(new Date()).subtract(24, 'hours').subtract('minutes', 1).toISOString();
                const today = new Date().toISOString();

                aggregation.bucketDefinition.ranges = [
                    {
                        to: pastYear
                    },
                    {
                        from: pastYear,
                        to: past3Months
                    },
                    {
                        from: past3Months,
                        to: pastMonth
                    },
                    {
                        from: pastMonth,
                        to: pastWeek
                    },
                    {
                        from: pastWeek,
                        to: past24hours
                    },
                    {
                        from: past24hours,
                        to: today
                    },
                    {
                        from: today
                    }
                ];
            }

            return aggregation;
        });

        // Build aggregation filters
        if (dataContext.filters.selectedFilters.length > 0) {

            // Make sure, if we have multiple filters, at least two filters have values to avoid apply an operator ('or','and') on only one condition failing the query.
            if (dataContext.filters.selectedFilters.length > 1 && dataContext.filters.selectedFilters.filter(selectedFilter => selectedFilter.values.length > 0).length > 1) {
                const refinementString = DataFilterHelper.buildFqlRefinementString(dataContext.filters.selectedFilters, dataContext.filters.filtersConfiguration, this.moment).join(',');
                if (!isEmpty(refinementString)) {
                    aggregationFilters = aggregationFilters.concat([`${dataContext.filters.filterOperator}(${refinementString})`]);
                }

            } else {
                aggregationFilters = aggregationFilters.concat(DataFilterHelper.buildFqlRefinementString(dataContext.filters.selectedFilters, dataContext.filters.filtersConfiguration, this.moment));
            }
        }

        if (this.properties.entityTypes.indexOf(EntityType.ExternalItem) !== -1) {
            // Build external connection ID
            this.properties.contentSourceConnectionIds.forEach(id => {
                contentSources.push(`/external/connections/${id}`);
            });
        }

        if (this.properties.entityTypes.indexOf(EntityType.ListItem) !== -1) {

            // Build sort properties (only relevant for SharePoint manged properties)
            this.properties.sortProperties.filter(s => s.sortField).forEach(sortProperty => {

                sortProperties.push({
                    name: sortProperty.sortField,
                    isDescending: sortProperty.sortDirection === SortFieldDirection.Descending ? true : false
                });
            });
        }

        // Build search request
        let searchRequest: IMicrosoftSearchRequest = {
            entityTypes: this.properties.entityTypes,
            query: {
                queryString: queryText
            },
            from: from,
            size: dataContext.itemsCountPerPage
        };

        if (this.properties.fields.length > 0) {
            searchRequest.fields = this.properties.fields.filter(a => a); // Fix to remove null values
        }

        if (aggregations.length > 0) {
            searchRequest.aggregations = aggregations.filter(a => a);
        }

        if (aggregationFilters.length > 0) {
            searchRequest.aggregationFilters = aggregationFilters;
        }

        if (sortProperties.length > 0) {
            searchRequest.sortProperties = sortProperties;
        }

        if (contentSources.length > 0) {
            searchRequest.contentSources = contentSources;
        }

        searchQuery.requests.push(searchRequest);

        return searchQuery;
    }

    /**
     * Retrieves data from Microsoft Graph API
     * @param searchRequest the Microsoft Search search request
     */
    private async search(searchQuery: IMicrosoftSearchQuery): Promise<IMicrosoftSearchDataSourceData> {

        const response: IMicrosoftSearchDataSourceData = await this._microsoftSearchService.search(this._microsoftSearchUrl, searchQuery, this.properties.useCustomAadApplication, this.properties.customAadApplicationOptions);
        this._itemsCount = this._microsoftSearchService.itemsCount;
        return response;
    }

    private parseAndCleanOptions(options: IComboBoxOption[]): IComboBoxOption[] {
        let optionWithComma = options.find(o => (o.key as string).indexOf(",") > 0);
        if (optionWithComma) {
            return (optionWithComma.key as string).split(",").map(k => { return { key: k.trim(), text: k.trim(), selected: true }; });
        }
        return options;
    }

    private async getAvailableProperties(): Promise<IComboBoxOption[]> {

        const searchManagedProperties = await this._sharePointSearchService.getAvailableManagedProperties();

        this._availableManagedProperties = searchManagedProperties.map(managedProperty => {
            return {
                key: managedProperty.name,
                text: managedProperty.name,
            } as IComboBoxOption;
        });

        return this._availableManagedProperties;
    }
}