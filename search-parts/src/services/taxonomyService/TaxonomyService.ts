import { Text, Log, Guid } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { ITaxonomyService } from './ITaxonomyService';
import { Constants } from '../../common/Constants';
import { ITerms, ITerm } from './ITaxonomyItems';
import { ITermGroupInfo, ITermSetInfo, ITermStoreInfo, ITermInfo } from '@pnp/sp/taxonomy';
import "@pnp/sp/taxonomy";
import { spfi, SPFI, SPFx } from '@pnp/sp';
import { PageContext } from '@microsoft/sp-page-context';

const TaxonomyService_ServiceKey = 'PnPModernSearchTaxonomyService';

export class TaxonomyService implements ITaxonomyService {

	public static ServiceKey: ServiceKey<ITaxonomyService> = ServiceKey.create(TaxonomyService_ServiceKey, TaxonomyService);

	/**
	 * The SPHttpClient instance
	 */
	private spHttpClient: SPHttpClient;

	private _sp: SPFI;

	constructor(serviceScope: ServiceScope) {

		serviceScope.whenFinished(() => {
			this.spHttpClient = serviceScope.consume<SPHttpClient>(SPHttpClient.serviceKey);
			const pageContext = serviceScope.consume<PageContext>(PageContext.serviceKey);
			this._sp = spfi().using(SPFx({
				pageContext: pageContext
			}));
		});
	}

	/**
	 * Gets multiple terms by their ids using the current taxonomy context
	 * @param siteUrl The site URL to use for the taxonomy session 
	 * @param termIds An array of term ids to search for
	 * @return {Promise<ITerm[]>} A promise containing the terms.
	 */
	public async getTermsById(siteUrl: string, termIds: string[]): Promise<ITerm[]> {

		let terms: ITerm[] = [];
		const clientServiceUrl = `${siteUrl}/_vti_bin/client.svc/ProcessQuery`;

		// Build XML query parameters
		const xmlQueryParameters = termIds.map(id => {
			return `<Object Type="String">${id}</Object>`;
		});

		const data = `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="pnp"><Actions><ObjectPath Id="1" ObjectPathId="0"/><ObjectPath Id="3" ObjectPathId="2"/><ObjectPath Id="5" ObjectPathId="4"/><Query Id="6" ObjectPathId="4"><Query SelectAllProperties="true"><Properties/></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="Id" SelectAll="true"/><Property Name="Labels" SelectAll="true"/></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="0" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}"/><Method Id="2" ParentId="0" Name="GetDefaultSiteCollectionTermStore"/><Method Id="4" ParentId="2" Name="GetTermsById"><Parameters><Parameter Type="Array">${xmlQueryParameters.join('')}</Parameter></Parameters></Method></ObjectPaths></Request>`;

		const response = await this.spHttpClient.post(clientServiceUrl, SPHttpClient.configurations.v1, {
			headers: {
				'Accept': 'application/json;odata.metadata=none',
				'X-ClientService-ClientTag': Constants.X_CLIENTSERVICE_CLIENTTAG,
				'UserAgent': Constants.X_CLIENTSERVICE_CLIENTTAG,
				'Content-Type': 'application/xml'
			},
			body: data
		});

		if (response.status === 200) {
			const responseJson = await response.json();

			// Retrieve the term collection results
			const termStoreResultTerms: ITerms[] = responseJson.filter((r: { [x: string]: string; }) => r['_ObjectType_'] === 'SP.Taxonomy.TermCollection');

			if (termStoreResultTerms.length > 0) {

				// Retrieve all terms
				terms = termStoreResultTerms[0]._Child_Items_;
			}
		} else {

			Log.error(TaxonomyService_ServiceKey, new Error(response.statusText));
			throw new Error(response.statusText);
		}

		return terms;
	}

	public async getTermGroups(): Promise<ITermGroupInfo[]> {
		const termGroups: ITermGroupInfo[] = await this._sp.termStore.groups();
		return termGroups;
	}

	public async getTermSets(groupId: string): Promise<ITermSetInfo[]> {
		const termSets: ITermSetInfo[] = await this._sp.termStore.groups.getById(groupId).sets();
		return termSets;
	}

	public async getTermStoreInfo(): Promise<ITermStoreInfo | undefined> {
		const termStoreInfo = await this._sp.termStore();
		return termStoreInfo;
	}

	public async getTermSetInfo(termSetId: Guid): Promise<ITermSetInfo | undefined> {
		const tsInfo = await this._sp.termStore.sets.getById(termSetId.toString())();
		return tsInfo;
	}

	public async getTermById(termSetId: Guid, termId: Guid): Promise<ITermInfo> {
		if (termId === Guid.empty) {
			return undefined;
		}
		try {
			const termInfo = await this._sp.termStore.sets.getById(termSetId.toString())
				.terms
				.getById(termId.toString())
				.select('id', 'labels', 'descriptions', 'properties', 'localProperties', 'ShortName', 'createdDateTime', 'lastModifiedDateTime', 'childrenCount', 'isAvailableForTagging', 'customSortOrder', 'isDeprecated', 'topicRequested')();
			return termInfo;
		} catch (error) {
			return undefined;
		}
	}

	public async getTermInfoById(termSetId: Guid, termId: Guid): Promise<ITermInfo> {
		if (termId === Guid.empty) {
			return undefined;
		}
		try {
			const termInfo = await this._sp.termStore.sets.getById(termSetId.toString()).terms.getById(termId.toString()).expand("parent")();
			return termInfo;
		} catch (error) {
			return undefined;
		}
	}
}