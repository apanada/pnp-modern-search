
export interface IAzureSearchQuery {
    searchTerm?: string;
    userInfo?: IUserInformation;
    pageInfo?: IPageContextInformation;
    pageNumber?: number;
    numberOfResultsPerPage?: number;
}

export interface IUserInformation {

    /**
     * The User identifier
     */
    UserId: string;
}

export interface IPageContextInformation {

    /**
     * The site absolute URL
     */
    siteAbsoluteUrl: string;

    /**
     *  The farm label
     */
    farmLabel: string;

    /**
     * The form digest value
     */
    formDigestValue: string;

    /**
     * The AAD tenant identifier
     */
    aadTenantId: string;

    /**
     * The AAD user identifier
     */
    aadUserId: string;
}