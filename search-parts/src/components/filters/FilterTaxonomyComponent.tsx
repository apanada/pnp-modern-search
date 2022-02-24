import * as React from 'react';
import { BaseWebComponent, FilterComparisonOperator, IDataFilterInfo, IDataFilterValueInfo, IDataFilterInternal, ExtensibilityConstants, IDataFilterValueInternal } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MessageBar, MessageBarType } from "office-ui-fabric-react";
import { ModernTaxonomyPicker } from 'shell-search-extensibility/lib/controls/modernTaxonomyPicker';
import { Guid, ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { ITermInfo, ITermStoreInfo, ITermSetInfo } from '@pnp/sp/taxonomy';
import "@pnp/sp/taxonomy";
import { dateAdd, PnPClientStorage } from '@pnp/common';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import { PageContext } from '@microsoft/sp-page-context';
import { TaxonomyService } from '../../services/taxonomyService/TaxonomyService';
import { ITaxonomyService } from '../../services/taxonomyService/ITaxonomyService';

export interface IFilterTaxonomyComponentProps {

    /**
     * The current selected filters. Because we can select values outside of values retrieved from results, we need this information to display the default date picker values correctly after the user selection
     */
    filter: IDataFilterInternal;

    /**
     * The current theme settings
     */
    themeVariant?: IReadonlyTheme;

    /**
     * Handler when the date is updated
     */
    onUpdate: (filterValues: IDataFilterValueInfo[]) => void;

    /**
     * The current page context reference
     */
    pageContext: PageContext;

    /**
     * The term store info
     */
    termStoreInfo?: ITermStoreInfo;

    /**
     * The termset info
     */
    termSetInfo?: ITermSetInfo;

    /**
     * The client storage instance
     */
    clientStorage: PnPClientStorage;

    /**
     * The taxonomy service instance
     */
    taxonomyService: ITaxonomyService;
}

export interface IFilterTaxonomyComponentState {

    /**
     * The current selected terms
     */
    selectedTerms: ITermInfo[];
}

export interface ITermDetails {
    id: string;
    name: string;
    selected: boolean;
}

export class FilterTaxonomyComponent extends React.Component<IFilterTaxonomyComponentProps, IFilterTaxonomyComponentState> {

    public constructor(props: IFilterTaxonomyComponentProps) {
        super(props);

        this.state = {
            selectedTerms: []
        };
    }

    public render() {

        let selectedTerms = this.state.selectedTerms;

        if (this.props.clientStorage.local.get(`${this.props.filter.filterName}-terms`)) {
            selectedTerms = this.props.clientStorage.local.get(`${this.props.filter.filterName}-terms`);
        }

        return <div>
            <ModernTaxonomyPicker
                key={`${this.props.filter.filterName}-taxonomy-filter`}
                allowMultipleSelections={true}
                termSetId={this.props.filter.termSetId}
                panelTitle={this.props.filter.displayName}
                label={this.props.filter.filterName}
                placeHolder="Search a value..."
                initialValues={selectedTerms}
                termStoreInfo={this.props.termStoreInfo}
                termSetInfo={this.props.termSetInfo}
                labelRequired={false}
                pageContext={this.props.pageContext}
                onChange={this._onModernPickerChange.bind(this)} />
        </div>;
    }

    public async componentDidMount() {
        this.props.clientStorage.local.deleteExpired();

        if (this.props.filter.values && this.props.filter.values.length > 0) {
            const initialValues = this._getInitialActiveFilterValues(this.props.filter.values);

            let data = await this._setInitialTerms(initialValues);

            this.setState({
                selectedTerms: data
            }, () => {
                if (this.state.selectedTerms && this.state.selectedTerms.length === 0) {
                    this.props.clientStorage.local.delete(`${this.props.filter.filterName}-terms`);
                    setTimeout(() => {
                        this._updateFilter(this.state.selectedTerms, true);
                    }, 500);
                } else {
                    this.props.clientStorage.local.put(`${this.props.filter.filterName}-terms`, this.state.selectedTerms, dateAdd(new Date(), 'day', 1));
                }
            });
        }
    }

    private _onModernPickerChange(terms: ITermInfo[], changeDetected: boolean) {

        this.setState({
            selectedTerms: terms
        }, () => {
            this.props.clientStorage.local.put(`${this.props.filter.filterName}-terms`, this.state.selectedTerms, dateAdd(new Date(), 'day', 1));
            this._updateFilter(this.state.selectedTerms, true);
        });
    }

    private _updateFilter(terms: ITermInfo[], selected: boolean) {

        let updatedValues: IDataFilterValueInfo[] = [];

        if (terms && terms.length > 0) {
            terms.map((term: ITermInfo) => {
                let termName = term && term.labels.length > 0 ? term.labels[0].name : null;

                // Build values
                if (termName) {
                    updatedValues.push({
                        name: termName,
                        value: term.id,
                        operator: FilterComparisonOperator.Eq,
                        selected: selected
                    });
                }
            });
        }

        this.props.onUpdate(updatedValues);
    }

    private _setInitialTerms = async (initialValues: ITermDetails[]): Promise<ITermInfo[]> => {
        if (Array.isArray(initialValues) && initialValues.length > 0) {
            var promises: Promise<ITermInfo>[] = initialValues.map(async (termDetails: ITermDetails) => {
                const term = await this.props.taxonomyService.getTermInfoById(Guid.parse(this.props.filter.termSetId), Guid.parse(termDetails.id));
                return new Promise<ITermInfo>((resolve, reject) => resolve(term));
            });

            var results: Promise<ITermInfo[]> = Promise.all(promises);
            const terms: ITermInfo[] = await results as ITermInfo[];

            let { selectedTerms } = this.state;
            const initialTermsState = selectedTerms ?? [];
            terms.map(v => initialTermsState.push({
                id: v.id,
                createdDateTime: v.createdDateTime,
                childrenCount: v.childrenCount,
                customSortOrder: v.customSortOrder,
                descriptions: v.descriptions,
                isAvailableForTagging: v.isAvailableForTagging,
                isDeprecated: v.isDeprecated,
                labels: v.labels,
                lastModifiedDateTime: v.lastModifiedDateTime,
                topicRequested: v.topicRequested,
                localProperties: v.localProperties,
                parent: v.parent,
                properties: v.properties
            }));

            return initialTermsState;
        }

        return [] as ITermInfo[];
    }

    private _getInitialActiveFilterValues = (initialFilterValues: IDataFilterValueInternal[]): ITermDetails[] => {
        let initialValues: ITermDetails[] = [];

        if (initialFilterValues && initialFilterValues.length > 0) {
            initialFilterValues.filter(value => value.selected).forEach(filterValue => {
                initialValues.push({
                    id: filterValue.value,
                    name: filterValue.name,
                    selected: filterValue.selected
                });
            });
        }

        return initialValues;
    }
}

export class FilterTaxonomyWebComponent extends BaseWebComponent {

    /**
     * The client storage instance
     */
    private clientStorage: PnPClientStorage;

    private _pageContext: PageContext;

    public constructor() {
        super();

        this.clientStorage = new PnPClientStorage();

        this._serviceScope.whenFinished(() => {

            const pageContext = this._serviceScope.consume<PageContext>(PageContext.serviceKey);
            this._pageContext = pageContext;
        });
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        let renderTaxonomyPicker: JSX.Element = null;
        let serviceScope: ServiceScope = this._serviceScope; // Default is the root shared service scope regardless the current Web Part 
        let taxonomyServiceKey: ServiceKey<any> = TaxonomyService.ServiceKey; // Defaut service key for TaxonomyService

        if (props.instanceId) {

            // Get the service scope and keys associated to the current Web Part displaying the component
            serviceScope = this._webPartServiceScopes.get(props.instanceId) ? this._webPartServiceScopes.get(props.instanceId) : serviceScope;
            taxonomyServiceKey = this._webPartServiceKeys.get(props.instanceId) ? this._webPartServiceKeys.get(props.instanceId).TaxonomyService : taxonomyServiceKey;
        }

        if (props.filter) {

            const taxonomyService = serviceScope.consume<ITaxonomyService>(taxonomyServiceKey);

            let termStoreInfo = null;
            if (this.clientStorage.local.get("termStoreInfo")) {
                termStoreInfo = this.clientStorage.local.get("termStoreInfo");
            } else {
                termStoreInfo = await taxonomyService.getTermStoreInfo();
                this.clientStorage.local.put("termStoreInfo", termStoreInfo, dateAdd(new Date(), 'day', 1));
            }

            let termSetInfo = null;
            if (this.clientStorage.local.get("termSetInfo")) {
                termSetInfo = this.clientStorage.local.get("termSetInfo");
            } else {
                termSetInfo = await taxonomyService.getTermSetInfo(props.filter.termSetId);
                this.clientStorage.local.put("termSetInfo", termSetInfo, dateAdd(new Date(), 'day', 1));
            }

            const filter = props.filter as IDataFilterInternal;
            renderTaxonomyPicker = <FilterTaxonomyComponent
                {...props}
                clientStorage={this.clientStorage}
                pageContext={this._pageContext}
                taxonomyService={taxonomyService}
                termStoreInfo={termStoreInfo}
                termSetInfo={termSetInfo}
                filter={filter} onUpdate={((filterValues: IDataFilterValueInfo[]) => {

                    // Unselect all previous values
                    const updatedValues = filter.values.map(value => {

                        // Exclude current selected values
                        if (filterValues.filter(filterValue => { return filterValue.value === value.value; }).length === 0) {
                            return {
                                name: value.name,
                                selected: false,
                                value: value.value,
                                operator: value.operator
                            } as IDataFilterValueInfo;
                        }
                    });

                    // Bubble event through the DOM
                    this.dispatchEvent(new CustomEvent(ExtensibilityConstants.EVENT_FILTER_UPDATED, {
                        detail: {
                            filterName: filter.filterName,
                            filterValues: filterValues.concat(updatedValues.filter(v => v)),
                            instanceId: props.instanceId
                        } as IDataFilterInfo,
                        bubbles: true,
                        cancelable: true
                    }));
                }).bind(this)}
            />;
        } else {
            renderTaxonomyPicker = <MessageBar
                messageBarType={MessageBarType.warning}
                isMultiline={false}>
                {`Component <pnp-filtertaxonomy> misconfigured. The HTML attribute 'filter' is missing.`}
            </MessageBar>;
        }

        ReactDOM.render(renderTaxonomyPicker, this);
    }
}