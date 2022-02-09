import { IMicrosoftSearchDataSourceData } from "../../models/search/IMicrosoftSearchDataSourceData";
import { IMicrosoftSearchQuery } from "../../models/search/IMicrosoftSearchRequest";

export interface IMicrosoftSearchService {

     /**
     * Performs a search query against SharePoint
     * @param searchQuery The search query in KQL format
     * @return The search results
     */
     search(microsoftSearchUrl: string, searchQuery: IMicrosoftSearchQuery): Promise<IMicrosoftSearchDataSourceData>;

     itemsCount: number;
}