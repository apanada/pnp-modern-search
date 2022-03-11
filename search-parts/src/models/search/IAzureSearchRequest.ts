
export interface IAzureSearchQuery {
    searchTerm?: string;
    userInfo?: IUserInformation;
    pageInfo?: any;
    pageNumber?: number;
    numberOfResultsPerPage?: number;
}

export interface IUserInformation {

    /**
     * The User identifier
     */
    UserId: string;
}