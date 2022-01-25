import { HttpClient } from "@microsoft/sp-http";
import { INlpRequest } from "../../models/search/INlpRequest";
import { INlpResponse } from "../../models/search/INlpResponse";
import { INlpService } from "./INlpService";
import { Log, ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";

const NlpService_ServiceKey = 'PnPModernSearchNlpService';

export class NlpService implements INlpService {

    public static ServiceKey: ServiceKey<INlpService> = ServiceKey.create(NlpService_ServiceKey, NlpService);

    /**
    * The current page context instance
    */
    private pageContext: PageContext;

    /**
     * The current service scope
     */
    private serviceScope: ServiceScope;

    /**
     * The SPHttpClient instance
     */
    private httpClient: HttpClient;

    /**
    * The Service Url
    */
    private serviceUrl: string;

    constructor(serviceScope: ServiceScope) {
        this.serviceScope = serviceScope;

        serviceScope.whenFinished(async () => {

            this.pageContext = serviceScope.consume<PageContext>(PageContext.serviceKey);
            this.httpClient = serviceScope.consume<HttpClient>(HttpClient.serviceKey);
        });
    }

    public setServiceUrl(value: string): void {
        this.serviceUrl = value;
    }

    /**
     * Interprets the user search query intents and return the optimized SharePoint query counterpart
     * @param rawQuery the user raw query input
     */
    public async enhanceSearchQuery(rawQuery: string, isStaging: boolean): Promise<INlpResponse> {

        const postData: string = JSON.stringify({
            rawQuery: rawQuery,
            uiLanguage: this.pageContext.cultureInfo.currentUICultureName.split("-")[0],
            isStaging: isStaging
        } as INlpRequest);

        // Make the call to the optimizer service
        const url = this.serviceUrl;

        try {

            const results = await this.httpClient.post(url, HttpClient.configurations.v1, {
                body: postData,
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-Type': 'application/json; charset=utf-8',
                    'Cache-Control': 'no-cache'
                }
            });

            const response: INlpResponse = await results.json();

            if (results.status === 200) {
                return response;
            } else {
                const error = JSON.stringify(response);
                Log.error(`[NlpService.enhanceSearchQuery()]: Error: '${error}' for url '${url}'`, new Error(error), this.serviceScope);
                throw new Error(error);
            }
        } catch (error) {
            const errorMessage = error ? error.message : `Failed to fetch URL '${url}'`;
            Log.error(`[NlpService.enhanceSearchQuery()]: Error: '${errorMessage}' for url '${url}'`, error, this.serviceScope);
            throw new Error(errorMessage);
        }
    }
}