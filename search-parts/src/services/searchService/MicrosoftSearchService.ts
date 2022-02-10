import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PnPClientStorage } from "@pnp/common/storage";
import { PageContext } from '@microsoft/sp-page-context';
import { IMicrosoftSearchService } from "./IMicrosoftSearchService";
import { HttpClient, HttpClientResponse, IHttpClientOptions, MSGraphClientFactory } from "@microsoft/sp-http";
import { ICustomAadApplicationOptions, IMicrosoftSearchQuery } from "../../models/search/IMicrosoftSearchRequest";
import { IMicrosoftSearchDataSourceData } from "../../models/search/IMicrosoftSearchDataSourceData";
import { FilterComparisonOperator, IDataFilterResult, IDataFilterResultValue } from "@pnp/modern-search-extensibility";
import { IMicrosoftSearchResponse, IMicrosoftSearchResultSet } from "../../models/search/IMicrosoftSearchResponse";
import { isEmpty } from "@microsoft/sp-lodash-subset";

const SearchService_ServiceKey = 'pnpSearchResults:MicrosoftSearchService';

const GRAPH_SCOPES = [
    "Calendars.Read",
    "Contacts.Read",
    "ExternalItem.Read.All",
    "Files.Read.All",
    "Mail.Read",
    "People.Read",
    "Sites.Read.All",
    "User.Read",
    "User.Read.All",
    "openid",
    "profile"
];

export class MicrosoftSearchService implements IMicrosoftSearchService {

    public static ServiceKey: ServiceKey<IMicrosoftSearchService> = ServiceKey.create(SearchService_ServiceKey, MicrosoftSearchService);

    /**
     * The current page context instance
     */
    private pageContext: PageContext;

    /**
     * The current service scope
     */
    private serviceScope: ServiceScope;

    /**
     * The MsalClient instance
     */
    private msalClient: any;

    /**
     * The client storage instance
     */
    private clientStorage: PnPClientStorage;

    /**
     * The data source items count
     */
    private _itemsCount: number = 0;

    constructor(serviceScope: ServiceScope) {

        this.serviceScope = serviceScope;

        this.clientStorage = new PnPClientStorage();

        serviceScope.whenFinished(async () => {

            this.pageContext = serviceScope.consume<PageContext>(PageContext.serviceKey);
        });
    }

    public set itemsCount(value: number) { this._itemsCount = value; }

    public get itemsCount(): number { return this._itemsCount; }

    public initializeMsalClient = async (useCustomAadApplication: boolean, customAadApplicationOptions: ICustomAadApplicationOptions) => {
        if (useCustomAadApplication && isEmpty(this.msalClient)) {
            const { MsalClient } = await import(
                /* webpackChunkName: '@pnp/msaljsclient' */
                '@pnp/msaljsclient'
            );

            // note we do not provide scopes here as the second parameter. We certainly could and will get a token
            // based on those scopes by making a call to getToken() without a param.
            this.msalClient = new MsalClient({
                auth: {
                    authority: `https://login.microsoftonline.com/${customAadApplicationOptions.tenantId}/`,
                    clientId: customAadApplicationOptions.clientId,
                    redirectUri: customAadApplicationOptions.redirectUrl ?? `${this.pageContext.web.absoluteUrl}/SitePages/Search.aspx`,
                },
            });
        }
    }

    /**
     * Retrieves data from Microsoft Graph API
     * @param searchQuery the Microsoft Search search request
     */
    public async search(microsoftSearchUrl: string, searchQuery: IMicrosoftSearchQuery, useCustomAadApplication: boolean, customAadApplicationOptions: ICustomAadApplicationOptions): Promise<IMicrosoftSearchDataSourceData> {

        let itemsCount = 0;
        let response: IMicrosoftSearchDataSourceData = {
            items: [],
            filters: []
        };
        let aggregationResults: IDataFilterResult[] = [];

        let jsonResponse: IMicrosoftSearchResponse;

        if (useCustomAadApplication && customAadApplicationOptions &&
            !isEmpty(customAadApplicationOptions.tenantId) &&
            !isEmpty(customAadApplicationOptions.clientId) &&
            !isEmpty(customAadApplicationOptions.redirectUrl)) {

            // Get an instance to the MsalClient
            await this.initializeMsalClient(useCustomAadApplication, customAadApplicationOptions);
            const graphAccessToken = await this.msalClient.getToken(GRAPH_SCOPES);

            const httpClient = this.serviceScope.consume<HttpClient>(HttpClient.serviceKey);

            const httpClientPostOptions: IHttpClientOptions = {
                headers: {
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                    "Authorization": "Bearer " + graphAccessToken,
                    'Cache-Control': 'no-cache'
                },
                body: JSON.stringify(searchQuery)
            };

            const httpResponse: HttpClientResponse = await httpClient.post(microsoftSearchUrl, HttpClient.configurations.v1, httpClientPostOptions);
            if (httpResponse) {

                const httpResponseJSON: IMicrosoftSearchResponse = await httpResponse.json();
                jsonResponse = httpResponseJSON;
            }

        } else {

            // Get an instance to the MSGraphClient
            const msGraphClientFactory = this.serviceScope.consume<MSGraphClientFactory>(MSGraphClientFactory.serviceKey);
            const msGraphClient = await msGraphClientFactory.getClient();
            const request = await msGraphClient.api(microsoftSearchUrl);

            jsonResponse = await request.post(searchQuery);
        }

        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {

            jsonResponse.value.forEach((value: IMicrosoftSearchResultSet) => {

                // Map results
                value.hitsContainers.forEach(hitContainer => {
                    itemsCount += hitContainer.total;

                    if (hitContainer.hits) {

                        const hits = hitContainer.hits.map(hit => {

                            if (hit.resource.fields) {

                                // Flatten 'fields' to be usable with the Search Fitler WP as refiners
                                Object.keys(hit.resource.fields).forEach(field => {
                                    hit[field] = hit.resource.fields[field];
                                });
                            }

                            return hit;
                        });

                        response.items = response.items.concat(hits);
                    }

                    if (hitContainer.aggregations) {

                        // Map refinement results
                        hitContainer.aggregations.forEach((aggregation) => {

                            let values: IDataFilterResultValue[] = [];
                            aggregation.buckets.forEach((bucket) => {
                                values.push({
                                    count: bucket.count,
                                    name: bucket.key,
                                    value: bucket.aggregationFilterToken,
                                    operator: FilterComparisonOperator.Contains
                                } as IDataFilterResultValue);
                            });

                            aggregationResults.push({
                                filterName: aggregation.field,
                                values: values
                            });
                        });

                        response.filters = aggregationResults;
                    }
                });
            });
        }

        if (jsonResponse?.queryAlterationResponse) {
            response.queryAlterationResponse = jsonResponse.queryAlterationResponse;
        }

        this._itemsCount = itemsCount;

        return response;
    }
}