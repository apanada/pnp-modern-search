import { INlpResponse } from "../../../../models/search/INlpResponse";

export interface INlpSpellCheckPanelProps {
    rawResponse: INlpResponse;

    rawQueryText: string;

    _onSpellCheckCallback: (enhancedQueryText: string) => void;
}