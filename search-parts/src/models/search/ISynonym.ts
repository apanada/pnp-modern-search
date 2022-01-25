export interface ISynonymTable {
    [key: string]: string[];
}

export interface ISynonymFieldConfiguration {
    Term: string;
    Synonyms: string;
    TwoWays: boolean;
}