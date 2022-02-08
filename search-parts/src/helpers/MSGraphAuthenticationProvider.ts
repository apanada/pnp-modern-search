import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { ServiceScope } from "@microsoft/sp-core-library";
import { AadHttpClientFactory, AadTokenProvider, HttpClient, HttpClientResponse, IHttpClientOptions } from "@microsoft/sp-http";

interface GraphOnBehalfOfRequest {
	tenantId: string;
	accessToken: string;
	graphScopes: string[];
	instance: string;
	clientId: string;
	clientSecret: string;
}

export class MSGraphAuthenticationProvider implements AuthenticationProvider {
	private _serviceScope: ServiceScope;
	private _tokenProvider: AadTokenProvider;
	private _aadResourceEndpoint: string;

	constructor(serviceScope: ServiceScope, resourceEndpoint: string, tokenProvider: AadTokenProvider) {
		this._serviceScope = serviceScope;
		this._tokenProvider = tokenProvider;
		this._aadResourceEndpoint = resourceEndpoint;
	}

	/**
	 * This method will get called before every request to the msgraph server
	 * This should return a Promise that resolves to an accessToken (in case of success) or rejects with error (in case of failure)
	 * Basically this method will contain the implementation for getting and refreshing accessTokens
	 */
	public async getAccessToken(): Promise<string> {

		if (this._serviceScope && this._aadResourceEndpoint) {
			// Get an instance to the AadHttpClientFactory
			const aadHttpClientFactory = this._serviceScope.consume<AadHttpClientFactory>(AadHttpClientFactory.serviceKey);
			const aadHttpClient = await aadHttpClientFactory.getClient(this._aadResourceEndpoint);			
			
			const accessToken: string = await this._tokenProvider.getToken(this._aadResourceEndpoint)
			const graphAccessToken: string = await this.getGraphAccessToken({
				accessToken: accessToken,
				clientId: "e5a0959e-a8fc-4db0-bc79-8ce90d1d1436",
				clientSecret: "zBF7Q~otZz0vIO0HzvlP9iAe2KB03~HgtHbEw",
				graphScopes: [
					"Calendars.Read",
					"Contacts.Read",
					"ExternalItem.Read.All",
					"Files.Read.All",
					"Mail.Read",
					"People.Read",
					"Sites.Read.All",
					"User.Read",
					"User.Read.All"
				],
				instance: "https://login.microsoftonline.com/",
				tenantId: "9d40d69a-0f98-48c0-97e2-3f5852ba4e67"
			});

			return graphAccessToken;
		}
	}

	private async getGraphAccessToken(graphOnBehalfOfRequest: GraphOnBehalfOfRequest): Promise<any> {
		// AAD Tenant ID
		const tenant = graphOnBehalfOfRequest.tenantId;

		// User AAD Access Token
		const token = graphOnBehalfOfRequest.accessToken;

		// Set scopes for current graph access token
		let scopes = [];
		if (graphOnBehalfOfRequest.graphScopes && graphOnBehalfOfRequest.graphScopes.length > 0) {
			graphOnBehalfOfRequest.graphScopes.forEach(graphScope => {
				scopes.push(`https://graph.microsoft.com/${graphScope}`);
			});
		}

		const url = graphOnBehalfOfRequest.instance + tenant + "/oauth2/v2.0/token";

		const params = {
			client_id: graphOnBehalfOfRequest.clientId,
			client_secret: graphOnBehalfOfRequest.clientSecret,
			grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
			assertion: token,
			requested_token_use: "on_behalf_of",
			scope: scopes.join(" ")
		};

		const httpClient = this._serviceScope.consume<HttpClient>(HttpClient.serviceKey);
		const httpClientPostOptions: IHttpClientOptions = {
			headers: {
				"Accept": "application/json",
				"Content-Type": "application/x-www-form-urlencoded"
			},
			body: new URLSearchParams(params).toString()
		};

		const httpResponse: HttpClientResponse = await httpClient.post(url, HttpClient.configurations.v1, httpClientPostOptions);
        if (httpResponse) {
			const httpResponseJSON: any = await httpResponse.json();			
			return httpResponseJSON;
		}		
	}
}