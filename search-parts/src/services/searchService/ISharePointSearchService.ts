import ISharePointManagedProperty from "../../models/search/ISharePointManagedProperty";
import { ISharePointSearchResults } from "../../models/search/ISharePointSearchResults";
import { ISharePointSearchQuery } from "../../models/search/ISharePointSearchQuery";
import { ISynonymTable } from "../../models/search/ISynonym";
import { IHubSite } from "../../models/common/ISIte";

export interface ISharePointSearchService {

     /**
     * Performs a search query against SharePoint
     * @param searchQuery The search query in KQL format
     * @return The search results
     */
     search(searchQuery: ISharePointSearchQuery): Promise<ISharePointSearchResults>;

     /**
     * Get available SharePoint search managed properties from the search schema
     */
     getAvailableManagedProperties(): Promise<ISharePointManagedProperty[]>;

     /**
     * Get all available languages for the search query
     */
     getAvailableQueryLanguages(): Promise<any>;

     /**
     * Determine if a SharePoint managed property is sortable
     * @param property the SharePoint managed property
     */
     validateSortableProperty(property: string): Promise<boolean>;

     /**
     * Retrieves search query suggestions
     * @param query the term to suggest from
     */
     suggest(query: string): Promise<string[]>;

     setSynonymTable(value: ISynonymTable): void;

     getHubSiteInfo(siteUrl: string, siteId: string): Promise<IHubSite>;
}