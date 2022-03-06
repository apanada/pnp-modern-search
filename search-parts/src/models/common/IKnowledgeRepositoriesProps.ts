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
    name: string;
    url: string;
    enabled: boolean;
}