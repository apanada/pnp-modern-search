import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PnPClientStorage } from "@pnp/common/storage";
import { PageContext } from '@microsoft/sp-page-context';
import { IMicrosoftSearchService } from "./IMicrosoftSearchService";
import { HttpClient, HttpClientResponse, IHttpClientOptions } from "@microsoft/sp-http";
import { IMicrosoftSearchQuery } from "../../models/search/IMicrosoftSearchRequest";
import { IMicrosoftSearchDataSourceData } from "../../models/search/IMicrosoftSearchDataSourceData";
import { FilterComparisonOperator, IDataFilterResult, IDataFilterResultValue } from "@pnp/modern-search-extensibility";
import { IMicrosoftSearchResponse, IMicrosoftSearchResultSet } from "../../models/search/IMicrosoftSearchResponse";

const SearchService_ServiceKey = 'pnpSearchResults:MicrosoftSearchService';

export class MicrosoftSearchService implements IMicrosoftSearchService {

    public static ServiceKey: ServiceKey<IMicrosoftSearchService> = ServiceKey.create(SearchService_ServiceKey, MicrosoftSearchService);

    /**
     * The current page context instance
     */
    private pageContext: PageContext;

    /**
     * The SharePoint search service endpoint REST URL
     */
    private searchEndpointUrl: string;

    /**
     * The current service scope
     */
    private serviceScope: ServiceScope;

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

            const { MsalClient } = await import(
                /* webpackChunkName: '@pnp/msaljsclient' */
                '@pnp/msaljsclient'
            );

            // note we do not provide scopes here as the second parameter. We certainly could and will get a token
            // based on those scopes by making a call to getToken() without a param.
            this.msalClient = new MsalClient({
                auth: {
                    authority: "https://login.microsoftonline.com/M365x083241.onmicrosoft.com/",
                    clientId: "e5a0959e-a8fc-4db0-bc79-8ce90d1d1436",
                    redirectUri: `${this.pageContext.web.absoluteUrl}/SitePages/Search.aspx`,
                },
            });

        });
    }

    public set itemsCount(value: number) { this._itemsCount = value; }

    public get itemsCount(): number { return this._itemsCount; }

    /**
     * Retrieves data from Microsoft Graph API
     * @param searchRequest the Microsoft Search search request
     */
    public async search(microsoftSearchUrl: string, searchQuery: IMicrosoftSearchQuery): Promise<IMicrosoftSearchDataSourceData> {

        let itemsCount = 0;
        let response: IMicrosoftSearchDataSourceData = {
            items: [],
            filters: []
        };
        let aggregationResults: IDataFilterResult[] = [];

        const graphAccessToken = await this.msalClient.getToken(["Calendars.Read","Contacts.Read","ExternalItem.Read.All","Files.Read.All","Mail.Read","People.Read","Sites.Read.All","User.Read","User.Read.All","openid","profile"]);

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

            const httpResponseJSON: any = await httpResponse.json();
            const jsonResponse: IMicrosoftSearchResponse = httpResponseJSON;

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
        }

        return response;
    }
}