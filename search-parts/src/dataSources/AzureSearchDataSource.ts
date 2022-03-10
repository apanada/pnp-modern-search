import { ServiceScope } from "@microsoft/sp-core-library";
import { isEmpty } from "@microsoft/sp-lodash-subset";
import { PageContext } from "@microsoft/sp-page-context";
import { IPropertyPaneField, IPropertyPaneGroup, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseDataSource, BuiltinTemplateSlots, FilterBehavior, IDataContext, IDataSourceData, ITemplateSlot, ITokenService, PagingBehavior } from "@pnp/modern-search-extensibility";
import * as commonStrings from 'CommonStrings';
import { IAzureSearchQuery } from "../models/search/IAzureSearchRequest";
import { IAzureSearchResults, IEbscoSearchResults } from "../models/search/IAzureSearchResults";
import { AzureSearchService } from "../services/searchService/AzureSearchService";
import { IAzureSearchService } from "../services/searchService/IAzureSearchService";
import { TokenService } from "../services/tokenService/TokenService";

export interface IAzureSearchDataSourceProperties {

    /**
     * The azure function app endpoint url
     */
    azureFunctionEndpoint: string;
}

export class AzureSearchDataSource extends BaseDataSource<IAzureSearchDataSourceProperties> {

    private _pageContext: PageContext;
    private _tokenService: ITokenService;
    private _azureSearchService: IAzureSearchService;

    private _propertyPaneWebPartInformation: any = null;

    /**
     * The data source items count
     */
    private _itemsCount: number = 0;

    private _propertyFieldCollectionData: any = null;
    private _customCollectionFieldType: any = null;

    public constructor(serviceScope: ServiceScope) {
        super(serviceScope);

        serviceScope.whenFinished(() => {
            this._pageContext = serviceScope.consume<PageContext>(PageContext.serviceKey);
            this._tokenService = serviceScope.consume<ITokenService>(TokenService.ServiceKey);
            this._azureSearchService = serviceScope.consume<IAzureSearchService>(AzureSearchService.ServiceKey);
        });
    }

    public async onInit(): Promise<void> {

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

    public async getData(dataContext: IDataContext): Promise<IDataSourceData> {

        let results: IDataSourceData = {
            items: []
        };

        // Ensuring azure function endpoint is set before launching a search
        if (!isEmpty(this.properties.azureFunctionEndpoint)) {
            const searchQuery = await this.buildAzureSearchQuery(dataContext);
            const response: IAzureSearchResults<IEbscoSearchResults> = await this.search(searchQuery);
            if (response) {
                results.items = response.results;
            }
        } else {
            // If no azure function endpoint set, manually set the results to prevent
            // having the previous search results items count displayed.
            this._itemsCount = 0;
        }

        return results;
    }

    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

        let commonFields: IPropertyPaneField<any>[] = [
            PropertyPaneTextField('dataSourceProperties.azureFunctionEndpoint', {
                label: commonStrings.DataSources.AzureSearch.AzureFunctionEndpointFieldLabel,
                placeholder: commonStrings.DataSources.AzureSearch.AzureFunctionEndpointPlaceholder,
                description: commonStrings.DataSources.AzureSearch.AzureFunctionEndpointFieldDescription
            }),
        ];

        let groupFields: IPropertyPaneField<any>[] = [
            ...commonFields
        ];

        return [
            {
                groupName: commonStrings.DataSources.AzureSearch.SourceConfigurationGroupName,
                groupFields: groupFields
            }
        ];
    }

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any) {
    }

    public onCustomPropertyUpdate(propertyPath: string, newValue: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
    }

    public getTemplateSlots(): ITemplateSlot[] {
        return [];
    }

    private async initProperties(): Promise<void> {
        this.properties.azureFunctionEndpoint = "https://graph-search-service.azurewebsites.net/api/GetEbscoResults?code=xpUpKujeoonybGU/pw8sjApmB1a1gcfvBOTADUUqgOrvWoketSNPKA==&clientId=apim-graph-search-service-apim"
    }

    private async buildAzureSearchQuery(dataContext: IDataContext): Promise<IAzureSearchQuery> {

        let searchQuery: IAzureSearchQuery = null;

        let queryText = '*'; // Default query string if not specified, the API does not support empty value
        let from = 0;

        // Query text
        if (dataContext.inputQueryText) {
            queryText = await this._tokenService.resolveTokens(dataContext.inputQueryText);
        }

        // Paging
        if (dataContext.pageNumber > 1) {
            from = (dataContext.pageNumber - 1) * dataContext.itemsCountPerPage;
        }

        // Build search query
        searchQuery = {
            searchTerm: queryText,
            pageNumber: from,
            numberOfResultsPerPage: dataContext.itemsCountPerPage,
            userInfo: {
                UserId: this._pageContext.legacyPageContext.aadUserId as string
            },
            pageInfo: {
                aadUserId: this._pageContext.legacyPageContext.aadUserId as string,
                aadTenantId: this._pageContext.legacyPageContext.aadTenantId as string,
                farmLabel: this._pageContext.legacyPageContext.farmLabel as string,
                formDigestValue: this._pageContext.legacyPageContext.formDigestValue as string,
                siteAbsoluteUrl: this._pageContext.site.absoluteUrl,
            }
        };

        return searchQuery;
    }

    /**
     * Retrieves data from Azure Function API
     * @param searchRequest the Azure Function API search request
     */
    private async search(searchQuery: IAzureSearchQuery): Promise<IAzureSearchResults<IEbscoSearchResults>> {

        const response: IAzureSearchResults<IEbscoSearchResults> = await this._azureSearchService.search(this.properties.azureFunctionEndpoint, searchQuery);
        this._itemsCount = this._azureSearchService.itemsCount;
        return response;
    }
}