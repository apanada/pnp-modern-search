import ISharePointManagedProperty from "../../models/search/ISharePointManagedProperty";
import { ISharePointSearchResults } from "../../models/search/ISharePointSearchResults";
import { ISharePointSearchQuery } from "../../models/search/ISharePointSearchQuery";
import { ISynonymTable } from "../../models/search/ISynonym";
import { IHubSite } from "../../models/common/ISIte";
import { IAccessRequest, IAccessRequestResults } from "../../models/common/IAccessRequest";

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

     checkUserAccessToReports(siteUrl: string, reportUrl: string): Promise<{ hasAccess: boolean, ItemCount: number }>;

     followDocument(documentUrl: string): Promise<boolean>;

     isDocumentFollowed(documentUrl: string): Promise<boolean>;

     stopFollowingDocument(documentUrl: string): Promise<boolean>;

     validateAccessRequest(listName: string, reportNumber: string): Promise<IAccessRequestResults | string>;

     submitAccessRequest(listName: string, accessRequest: IAccessRequest): Promise<IAccessRequestResults | string>;
}