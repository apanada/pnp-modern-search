import { INlpResponse } from "../../../models/search/INlpResponse";

export interface ISearchBoxContainerState {

    /**
     * The current value of the input string
     */
    searchInputValue: string;

    /**
     * Error message
     */
    errorMessage: string;

    /**
     * Show Clear button in the Search Box
     */
    showClearButton: boolean;

    /**
     * The enhanced query response
     */
    enhancedQuery: INlpResponse;

    showSpellCheckPanel: boolean;
}