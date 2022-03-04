export interface INlpRequest {

    /**
     * The raw query from the user in the search box
     */
    rawQuery: string;

    /**
     * The current UI language. Used to determine the language for optimization
     */
    uiLanguage: string;
}