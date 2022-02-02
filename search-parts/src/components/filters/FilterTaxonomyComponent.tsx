import * as React from 'react';
import { BaseWebComponent, FilterComparisonOperator, IDataFilterInfo, IDataFilterValueInfo, IDataFilterInternal, ExtensibilityConstants, IDataFilterValueInternal } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ITheme } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Link, MessageBar, MessageBarType } from "office-ui-fabric-react";
import * as strings from 'CommonStrings';
import { ModernTaxonomyPicker } from 'shell-search-extensibility/lib/controls/modernTaxonomyPicker';
import { Guid, ServiceScope } from '@microsoft/sp-core-library';
import { ITermInfo, ITermStoreInfo } from '@pnp/sp/taxonomy';
import { sp } from "shell-search-extensibility/lib/index";
import "@pnp/sp/taxonomy";

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
     * The current service scope reference
     */
    serviceScope: ServiceScope;

    termStoreInfo?: ITermStoreInfo;
}

export interface IFilterTaxonomyComponentState {

    /**
     * The current selected terms
     */
    selectedTerms: ITermInfo[];

    initialFilterValues: IDataFilterValueInternal[];
}

export interface ITermDetails {
    id: string;
    name: string;
    selected: boolean;
}

const TERM_SET_ID: string = "8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f";

export class FilterTaxonomyComponent extends React.Component<IFilterTaxonomyComponentProps, IFilterTaxonomyComponentState> {

    public constructor(props: IFilterTaxonomyComponentProps) {
        super(props);

        this.state = {
            selectedTerms: null,
            initialFilterValues: null
        };
    }

    public render() {

        if (this.state.selectedTerms && this.state.selectedTerms.length > 0) {
            return <div>
                <ModernTaxonomyPicker allowMultipleSelections={true}
                    termSetId={TERM_SET_ID}
                    panelTitle="Departments"
                    label="Departments"
                    placeHolder="Search a value..."
                    initialValues={this.state.selectedTerms}
                    termStoreInfo={this.props.termStoreInfo}
                    labelRequired={false}
                    serviceScope={this.props.serviceScope}
                    onChange={this._onModernPickerChange.bind(this)} />
            </div>;
        } else {
            return <div>
                <ModernTaxonomyPicker allowMultipleSelections={true}
                    termSetId={TERM_SET_ID}
                    panelTitle="Departments"
                    label="Departments"
                    placeHolder="Search a value..."
                    termStoreInfo={this.props.termStoreInfo}
                    labelRequired={false}
                    serviceScope={this.props.serviceScope}
                    onChange={this._onPickerChange.bind(this)} />
            </div>;
        }
    }

    public componentDidMount() {
        this.setState({
            initialFilterValues: this.props.filter.values
        }, async () => {
            if (this.state.initialFilterValues && this.state.initialFilterValues.length > 0) {
                const initialValues = this._getInitialActiveFilterValues(this.state.initialFilterValues);

                let data = await this._setInitialTerms(initialValues);

                this.setState({
                    selectedTerms: data
                }, () => {
                    console.log(this.state);
                    if (this.state.selectedTerms && this.state.selectedTerms.length > 0) {
                        this._updateFilter(this.state.selectedTerms, true);
                    }
                });
            }
        });
    }

    private _onModernPickerChange(terms: ITermInfo[]) {

        if (this.state.selectedTerms && this.state.selectedTerms.length > 0) {
            this.setState({
                selectedTerms: terms
            }, () => {
                if (terms && terms.length === 0) {
                    localStorage.setItem("emptyPicker", JSON.stringify(false));
                }
                this._updateFilter(this.state.selectedTerms, true);
            });
        }
    }

    private _onPickerChange(terms: ITermInfo[]) {

        if (this.state.selectedTerms === null || (this.state.selectedTerms && this.state.selectedTerms.length === 0)) {
            this.setState({
                selectedTerms: terms
            }, () => {
                const initialValues = this._getInitialActiveFilterValues(this.state.initialFilterValues);
                if (terms && terms.length > 0 && initialValues && initialValues.length === 0) {
                    if (!localStorage.getItem("emptyPicker") || (localStorage.getItem("emptyPicker") && JSON.parse(localStorage.getItem("emptyPicker")) === false)) {
                        localStorage.setItem("emptyPicker", JSON.stringify(true));
                        this._updateFilter(this.state.selectedTerms, true);
                    }
                }
            });
        }
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

    private async getTermById(termSetId: Guid, termId: Guid): Promise<ITermInfo> {
        if (termId === Guid.empty) {
            return undefined;
        }
        try {
            const termInfo = await sp.termStore.sets.getById(termSetId.toString()).terms.getById(termId.toString()).expand("parent")();
            return termInfo;
        } catch (error) {
            return undefined;
        }
    }

    private _setInitialTerms = async (initialValues: ITermDetails[]): Promise<ITermInfo[]> => {
        if (Array.isArray(initialValues) && initialValues.length > 0) {
            var promises: Promise<ITermInfo>[] = initialValues.map(async (termDetails: ITermDetails) => {
                const term = await this.getTermById(Guid.parse(TERM_SET_ID), Guid.parse(termDetails.id));
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

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        let renderTaxonomyPicker: JSX.Element = null;

        if (props.filter) {

            let termStoreInfo = null;
            if (localStorage.getItem("termStoreInfo")) {
                termStoreInfo = JSON.parse(localStorage.getItem("termStoreInfo"));
            } else {
                termStoreInfo = await this.getTermStoreInfo();
                localStorage.setItem("termStoreInfo", JSON.stringify(termStoreInfo));
            }

            const filter = props.filter as IDataFilterInternal;
            renderTaxonomyPicker = <FilterTaxonomyComponent {...props} serviceScope={this._serviceScope} termStoreInfo={termStoreInfo} filter={filter} onUpdate={((filterValues: IDataFilterValueInfo[]) => {

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

    public async getTermStoreInfo(): Promise<ITermStoreInfo | undefined> {
        const termStoreInfo = await sp.termStore();
        return termStoreInfo;
    }
}