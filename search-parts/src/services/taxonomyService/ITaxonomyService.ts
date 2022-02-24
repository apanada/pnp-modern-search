import { Guid } from '@microsoft/sp-core-library';
import { ITermGroupInfo, ITermSetInfo, ITermStoreInfo, ITermInfo } from '@pnp/sp/taxonomy';
import { ITerm } from './ITaxonomyItems';

export interface ITaxonomyService {

    /**
     * Gets multiple terms by their ids using the current taxonomy context
     * @param siteUrl The site URL to use for the taxonomy session 
     * @param termIds An array of term ids to search for
     * @return {Promise<ITerm[]>} A promise containing the terms.
     */
    getTermsById(siteUrl: string, termIds: string[]): Promise<ITerm[]>;

    getTermGroups(): Promise<ITermGroupInfo[]>;

    getTermSets(groupId: string): Promise<ITermSetInfo[]>;

    getTermStoreInfo(): Promise<ITermStoreInfo | undefined>;

    getTermSetInfo(termSetId: Guid): Promise<ITermSetInfo | undefined>;

    getTermById(termSetId: Guid, termId: Guid): Promise<ITermInfo>;

    getTermInfoById(termSetId: Guid, termId: Guid): Promise<ITermInfo>;
}