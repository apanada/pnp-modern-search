import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { HttpClient, HttpClientResponse, IHttpClientOptions } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { IAzureSearchQuery } from "../../models/search/IAzureSearchRequest";
import { IAzureSearchResults, IEbscoSearchResults } from "../../models/search/IAzureSearchResults";
import { IAzureSearchService } from "./IAzureSearchService";

const SearchService_ServiceKey = 'pnpSearchResults:AzureSearchService';

export class AzureSearchService implements IAzureSearchService {

    public static ServiceKey: ServiceKey<AzureSearchService> = ServiceKey.create(SearchService_ServiceKey, AzureSearchService);

    /**
     * The current page context instance
     */
    private pageContext: PageContext;

    /**
     * The current service scope
     */
    private serviceScope: ServiceScope;

    /**
     * The HttpClient instance
     */
    private httpClient: HttpClient;

    /**
     * The data source items count
     */
    private _itemsCount: number = 0;

    constructor(serviceScope: ServiceScope) {

        this.serviceScope = serviceScope;

        serviceScope.whenFinished(async () => {

            this.pageContext = serviceScope.consume<PageContext>(PageContext.serviceKey);
            this.httpClient = serviceScope.consume<HttpClient>(HttpClient.serviceKey);
        });
    }

    public set itemsCount(value: number) { this._itemsCount = value; }

    public get itemsCount(): number { return this._itemsCount; }


    /**
     * Retrieves data from Microsoft Graph API
     * @param searchQuery the Microsoft Search search request
     */
    public async search(azureFunctionEndpointUrl: string, searchQuery: IAzureSearchQuery): Promise<IAzureSearchResults<IEbscoSearchResults>> {
        let response: IAzureSearchResults<IEbscoSearchResults> = {
            results: []
        };

        const httpClientPostOptions: IHttpClientOptions = {
            headers: {
                "Content-Type": "application/json",
                "Accept": "application/json",
                'Cache-Control': 'no-cache'
            },
            body: JSON.stringify(searchQuery)
        };

        const httpResponse: HttpClientResponse = await this.httpClient.post(azureFunctionEndpointUrl, HttpClient.configurations.v1, httpClientPostOptions);
        if (httpResponse) {

            const httpResponseJSON = await httpResponse.json();

            if (httpResponseJSON && httpResponseJSON.results && Array.isArray(httpResponseJSON.results) && httpResponseJSON.results.length > 0) {
                response.results.push(...httpResponseJSON.results as Array<IEbscoSearchResults>);
                this._itemsCount = response.results.length;
            }
        }

        return response;
    }
}