import { INlpResponse } from "../../models/search/INlpResponse";

export interface INlpService {

    /**
     * Interprets the user search query intents and return the relevant keywords
     * @param rawQuery the user raw query input
     * @param isStaging indicates if we should use the LUIS staging model 
     */
    enhanceSearchQuery(rawQuery: string, isStaging: boolean): Promise<INlpResponse>;

    setServiceUrl(value: string): void;
}