export interface IKnowledgeRepositoriesProps {
    /**
     * Flag indicating if knowledge repositories should be applied/enabled
     */
    knowledgeRepositoriesEnabled: boolean;

    /**
     * Knowledge repositories list
     */
    knowledgeRepositoriesList: IKnowledgeRepositoriesFieldConfiguration[];
}

export interface IKnowledgeRepositoriesFieldConfiguration {
    Name: string;
    Url: string;
    Enabled: boolean;
}