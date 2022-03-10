export interface IAzureSearchResults<T> {
    results?: Array<T>;
}

export interface IEbscoSearchResults {
    resultId?: string;
    pLink?: string;
    header?: {
        pubType?: string;
        pubTypeId?: string;
    },
    fullText?: {
        customLinks?: Array<{
            url?: string;
            text?: string;
            mouseOverText?: string;
            icon?: string;
        }>;
    }
    items?: Array<{
        name?: string;
        label?: string;
        group?: string;
        data?: string;
    }>;
    recordInfo?: {
        bibRecord?: {
            bibRelationships?: {
                isPartOfRelationships?: Array<{
                    bibEntity?: {
                        dates?: Array<{
                            d: number,
                            m: number,
                            y: number,
                            type: string
                        }>;
                    }
                }>;
            }
        }
    }
}