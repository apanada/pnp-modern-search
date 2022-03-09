import { IAzureSearchQuery } from "../../models/search/IAzureSearchRequest";
import { IAzureSearchResults, IEbscoSearchResults } from "../../models/search/IAzureSearchResults";

export interface IAzureSearchService {

    /**
    * Performs a search query against Azure Function API
    * @param searchQuery The search query
    * @return The search results
    */
    search(azureFunctionEndpointUrl: string, searchQuery: IAzureSearchQuery): Promise<IAzureSearchResults<IEbscoSearchResults>>;

    itemsCount: number;
}