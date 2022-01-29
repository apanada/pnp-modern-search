import * as React from 'react';
import { BaseWebComponent, FilterComparisonOperator, IDataFilterInfo, IDataFilterValueInfo, IDataFilterInternal, ExtensibilityConstants } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ITheme } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Link, MessageBar, MessageBarType } from "office-ui-fabric-react";
import * as strings from 'CommonStrings';
import { ModernTaxonomyPicker } from 'shell-search-extensibility/lib/controls/modernTaxonomyPicker';
import { Guid, ServiceScope } from '@microsoft/sp-core-library';
import { ITermInfo } from '@pnp/sp/taxonomy';
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
}

export interface IFilterTaxonomyComponentState {

    /**
     * The current selected terms
     */
    selectedTerms: ITermInfo[];

    initialTermsState: ITermInfo[];
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
            initialTermsState: null
        };
    }

    public render() {

        console.log(this.state.initialTermsState);

        return <div>
            <ModernTaxonomyPicker allowMultipleSelections={true}
                termSetId={TERM_SET_ID}
                panelTitle="Departments"
                label="Departments"
                placeHolder="Search a value..."
                initialValues={this.state.initialTermsState}
                serviceScope={this.props.serviceScope}
                onChange={this._onPickerChange.bind(this)} />
            <Link theme={this.props.themeVariant as ITheme}>{strings.Filters.ClearAllFiltersButtonLabel}</Link>
        </div>;
    }

    public componentDidMount() {

        if (this.props.filter.values.length > 0) {

            let initialValues: ITermDetails[] = [];

            this.props.filter.values.filter(value => value.selected).forEach(filterValue => {
                initialValues.push({
                    id: filterValue.value,
                    name: filterValue.name,
                    selected: filterValue.selected
                });
            });

            if (Array.isArray(initialValues) && initialValues.length > 0) {
                var promises: Promise<ITermInfo>[] = initialValues.map(async (termDetails: ITermDetails) => {
                    const term = await this.getTermById(Guid.parse(TERM_SET_ID), Guid.parse(termDetails.id));
                    return new Promise<ITermInfo>((resolve, reject) => resolve(term));
                });

                var results: Promise<ITermInfo[]> = Promise.all(promises);
                results.then((data: ITermInfo[]) => {
                    this.setState({ initialTermsState: data }, () => {
                        console.log("logged");
                    });
                });
            }
        }
    }

    private _onPickerChange(terms: ITermInfo[]) {

        this.setState({
            selectedTerms: terms
        });

        this._updateFilter(terms, true);
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
}

export class FilterTaxonomyWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        let renderTaxonomyPicker: JSX.Element = null;

        if (props.filter) {

            const filter = props.filter as IDataFilterInternal;
            renderTaxonomyPicker = <FilterTaxonomyComponent {...props} serviceScope={this._serviceScope} filter={filter} onUpdate={((filterValues: IDataFilterValueInfo[]) => {

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