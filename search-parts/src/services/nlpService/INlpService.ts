import { INlpResponse } from "../../models/search/INlpResponse";

export interface INlpService {

    /**
     * Interprets the user search query intents and return the relevant keywords
     * @param rawQuery the user raw query input
     */
    enhanceSearchQuery(rawQuery: string): Promise<INlpResponse>;

    setServiceUrl(value: string): void;
}