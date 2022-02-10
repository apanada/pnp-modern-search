import { IMicrosoftSearchDataSourceData } from "../../models/search/IMicrosoftSearchDataSourceData";
import { ICustomAadApplicationOptions, IMicrosoftSearchQuery } from "../../models/search/IMicrosoftSearchRequest";

export interface IMicrosoftSearchService {

     /**
     * Performs a search query against SharePoint
     * @param searchQuery The search query in KQL format
     * @return The search results
     */
     search(microsoftSearchUrl: string, searchQuery: IMicrosoftSearchQuery, useCustomAadApplication: boolean, customAadApplicationOptions: ICustomAadApplicationOptions): Promise<IMicrosoftSearchDataSourceData>;

     itemsCount: number;
}