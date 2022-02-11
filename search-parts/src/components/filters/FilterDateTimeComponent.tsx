import * as React from 'react';
import { BaseWebComponent, ExtensibilityConstants, FilterComparisonOperator, IDataFilterInfo, IDataFilterInternal, IDataFilterValueInfo } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ContextualMenuItemType, DefaultButton, IChoiceGroupOption, IContextualMenuItem, IContextualMenuListProps, IRenderFunction, ITheme, PrimaryButton } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IDatePickerProps, IDatePickerStyles, IDatePickerStyleProps, DatePicker, Link, MessageBar, MessageBarType } from "office-ui-fabric-react";
import { DateHelper } from '../../helpers/DateHelper';
import * as strings from 'CommonStrings';
import { DateFilterInterval } from './FilterDateIntervalComponent';
import styles from './FilterDateTimeComponent.module.scss';
import { isEmpty } from '@microsoft/sp-lodash-subset';

export interface IFilterDateTimeComponentProps {

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
     * The moment.js library reference
     */
    moment: any;
}

export interface IFilterDateTimeComponentState {

    /**
     * The current selected 'From' date
     */
    selectedFromDate: Date;

    /**
     * The current selected 'To' date
     */
    selectedToDate: Date;

    options: IContextualMenuItem[];

    selectedOptions: { [key: string]: boolean };

    enableApplyDate: boolean;
}

export class FilterDateTimeComponent extends React.Component<IFilterDateTimeComponentProps, IFilterDateTimeComponentState> {

    private _allOptions: IContextualMenuItem[];

    public constructor(props: IFilterDateTimeComponentProps) {
        super(props);

        this.state = {
            selectedFromDate: null,
            selectedToDate: null,
            options: [],
            selectedOptions: {},
            enableApplyDate: false
        };

        this._onToggleSelect = this._onToggleSelect.bind(this);
        this._updateFromDate = this._updateFromDate.bind(this);
        this._updateToDate = this._updateToDate.bind(this);
        this._updateFilter = this._updateFilter.bind(this);
        this._onFormatDate = this._onFormatDate.bind(this);
    }

    public render() {

        let dateAsString: string = undefined;
        const values = this.props.filter.values.filter(value => value.selected).sort((a, b) => {
            return new Date(a.value).getTime() - new Date(b.value).getTime();
        });

        // Values currently in the range (we take on the first one to determine the correct interval)
        if (values.length >= 1) {
            dateAsString = values[0].value;
        }

        return (
            <>
                <div className={styles.filterDatTime}>
                    <DefaultButton
                        text={this.props.filter.displayName}
                        menuProps={{
                            shouldFocusOnMount: true,
                            items: [
                                ...this.state.options,
                                { key: 'divider_1', itemType: ContextualMenuItemType.Divider }
                            ],
                            onRenderMenuList: this._renderMenuList.bind(this),
                        }}
                        styles={{ root: { border: "0px" } }} />
                </div>
            </>
        );
    }

    public componentDidMount() {

        const selectedOptions: {
            /**
             * The key represents the DateFilterInterval and the number equals to the checked status for this option
             */
            [key: string]: boolean
        } = {};

        this._allOptions = [
            {
                key: DateFilterInterval.AnyTime.toString(),
                text: strings.General.DateIntervalStrings.AnyTime,
                canCheck: true,
                checked: this.state.selectedOptions[DateFilterInterval.AnyTime.toString()] ?? false,
                onClick: this._onToggleSelect
            },
            {
                key: DateFilterInterval.Past24.toString(),
                text: strings.General.DateIntervalStrings.PastDay,
                canCheck: true,
                checked: this.state.selectedOptions[DateFilterInterval.Past24.toString()] ?? false,
                onClick: this._onToggleSelect
            },
            {
                key: DateFilterInterval.PastWeek.toString(),
                text: strings.General.DateIntervalStrings.PastWeek,
                canCheck: true,
                checked: this.state.selectedOptions[DateFilterInterval.PastWeek.toString()] ?? false,
                onClick: this._onToggleSelect
            },
            {
                key: DateFilterInterval.PastMonth.toString(),
                text: strings.General.DateIntervalStrings.PastMonth,
                canCheck: true,
                checked: this.state.selectedOptions[DateFilterInterval.PastMonth.toString()] ?? false,
                onClick: this._onToggleSelect
            },
            {
                key: DateFilterInterval.Past3Months.toString(),
                text: strings.General.DateIntervalStrings.Past3Months,
                canCheck: true,
                checked: this.state.selectedOptions[DateFilterInterval.Past3Months.toString()] ?? false,
                onClick: this._onToggleSelect
            },
            {
                key: DateFilterInterval.PastYear.toString(),
                text: strings.General.DateIntervalStrings.PastYear,
                canCheck: true,
                checked: this.state.selectedOptions[DateFilterInterval.PastYear.toString()] ?? false,
                onClick: this._onToggleSelect
            },
            {
                key: DateFilterInterval.OlderThanAYear.toString(),
                text: strings.General.DateIntervalStrings.Older,
                canCheck: true,
                checked: this.state.selectedOptions[DateFilterInterval.OlderThanAYear.toString()] ?? false,
                onClick: this._onToggleSelect
            }
        ];

        // Determine intervals according current filter values
        if (this.props.filter.values.length > 0) {
            let selectedFromDate: Date = undefined;
            let selectedToDate: Date = undefined;
            let availableOptions: IContextualMenuItem[] = this._allOptions ?? [];

            // Determine 'from' and 'to' dates by lokking at the operator for currently selected values
            this.props.filter.values.filter(value => value.selected).forEach(filterValue => {
                if (filterValue.operator === FilterComparisonOperator.Geq && !selectedFromDate) {
                    selectedFromDate = new Date(filterValue.value);
                }

                if ((filterValue.operator === FilterComparisonOperator.Lt || filterValue.operator === FilterComparisonOperator.Leq) && !selectedToDate) {
                    selectedToDate = new Date(filterValue.value);
                }
            });

            this.props.filter.values.forEach(value => {

                // Could have count 0 with SharePoint date ranges
                if (value.count > 0) {
                    const interval = this._getIntervalKeyForValue(value.value);
                    if (interval) {
                        if (Object.keys(selectedOptions).indexOf(interval) === -1) {
                            selectedOptions[interval] = value.selected;
                        } else {
                            selectedOptions[interval] = (selectedOptions[interval] || value.selected);
                        }
                    }
                }
            });

            availableOptions = this._allOptions.map(option => {

                if (Object.keys(selectedOptions).indexOf(option.key) !== -1) {
                    option.text = option.text;
                    option.checked = selectedOptions[option.key];
                    return option;
                } else if (option.key === DateFilterInterval.AnyTime.toString()) {
                    return option;
                }

            }).filter(o => o);

            this.setState({
                selectedFromDate: selectedFromDate,
                selectedToDate: selectedToDate,
                options: availableOptions,
                enableApplyDate: (selectedFromDate !== undefined) || (selectedToDate !== undefined)
            });
        }
    }

    private _onToggleSelect(ev?: React.MouseEvent<HTMLButtonElement>, item?: IContextualMenuItem) {
        ev && ev.preventDefault();
        if (item) {
            let options: IContextualMenuItem[] = this.state.options.map(option => {
                if (option.key === item.key) {
                    option.checked = true;
                }
                return option;
            });
            this.setState({
                selectedOptions: { ...this.state.selectedOptions, [item.key]: this.state.selectedOptions[item.key] === undefined ? true : !this.state.selectedOptions[item.key] },
                options: options
            });

            // Buld filters
            let updatedValues: IDataFilterValueInfo[] = [];

            switch (item.key) {

                case String(DateFilterInterval.OlderThanAYear):
                    updatedValues.push(
                        {
                            name: strings.General.DateIntervalStrings.Older,
                            value: this.props.moment(new Date()).subtract(1, 'years').subtract('minutes', 1).toISOString(), // Needed to distinguish past yeart VS older than a year
                            selected: true,
                            operator: FilterComparisonOperator.Lt
                        }
                    );
                    break;

                case String(DateFilterInterval.Past24):
                    updatedValues.push(
                        {
                            name: strings.General.DateIntervalStrings.PastDay,
                            value: new Date().toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Leq
                        },
                        {
                            name: strings.General.DateIntervalStrings.PastDay,
                            value: this.props.moment(new Date()).subtract(24, 'hours').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Geq
                        }
                    );
                    break;

                case String(DateFilterInterval.Past3Months):
                    updatedValues.push(
                        {
                            name: strings.General.DateIntervalStrings.Past3Months,
                            value: this.props.moment(new Date()).subtract(1, 'months').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Leq
                        },
                        {
                            name: strings.General.DateIntervalStrings.Past3Months,
                            value: this.props.moment(new Date()).subtract(3, 'months').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Geq
                        }
                    );
                    break;

                case String(DateFilterInterval.PastMonth):
                    updatedValues.push(
                        {
                            name: strings.General.DateIntervalStrings.PastMonth,
                            value: this.props.moment(new Date()).subtract(1, 'week').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Leq
                        },
                        {
                            name: strings.General.DateIntervalStrings.PastMonth,
                            value: this.props.moment(new Date()).subtract(1, 'months').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Geq
                        }
                    );
                    break;

                case String(DateFilterInterval.PastWeek):
                    updatedValues.push(
                        {
                            name: strings.General.DateIntervalStrings.PastWeek,
                            value: this.props.moment(new Date()).subtract(24, 'hours').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Leq
                        },
                        {
                            name: strings.General.DateIntervalStrings.PastWeek,
                            value: this.props.moment(new Date()).subtract(1, 'week').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Geq
                        }
                    );
                    break;

                case String(DateFilterInterval.PastYear):
                    updatedValues.push(
                        {
                            name: strings.General.DateIntervalStrings.PastYear,
                            value: this.props.moment(new Date()).subtract(3, 'months').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Leq
                        },
                        {
                            name: strings.General.DateIntervalStrings.PastYear,
                            value: this.props.moment(new Date()).subtract(1, 'years').toISOString(),
                            selected: true,
                            operator: FilterComparisonOperator.Geq
                        }
                    );
                    break;
            }

            this.props.onUpdate(updatedValues);
        }
    }

    private _getIntervalDate(unit: string, count: number): Date {
        return this._getIntervalDateFromStartDate(new Date(), unit, count);
    }

    private _getIntervalDateFromStartDate(startDate: Date, unit: string, count: number): Date {
        return this.props.moment(startDate).subtract(count, unit).toDate();
    }

    private _getIntervalKeyForValue(dateAsString: string): string {

        if (dateAsString) {

            // Value from SharePoint Search (RefinableDateXX properties)
            if (dateAsString.indexOf('range(') !== -1) {
                const matches = dateAsString.match(/(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}(?:\.\d*)?)((-(\d{2}):(\d{2})|Z)?)/gi);
                if (matches) {

                    // Return the last date of the range expression to get the correct interval 
                    dateAsString = matches[matches.length - 1];
                }
            }

            const selectedStartDate = new Date(dateAsString);

            // The 1 minute time addition represents a 'buffer' between the time the filter is selected in the component and the time when the filter value is sent back from the data source to the component 
            // If the data source takes too long to execute, it may cause of the wrong interval to be selected at render
            const pastYearDate = this._getIntervalDateFromStartDate(this._getIntervalDate("years", 1), 'minutes', 1);

            if (selectedStartDate < pastYearDate) {
                return DateFilterInterval.OlderThanAYear.toString();
            } else {

                const past24Date = this._getIntervalDateFromStartDate(this._getIntervalDate("days", 1), 'minutes', 1);
                const pastWeekDate = this._getIntervalDateFromStartDate(this._getIntervalDate("weeks", 1), 'minutes', 1);
                const pastMonthDate = this._getIntervalDateFromStartDate(this._getIntervalDate("months", 1), 'minutes', 1);
                const past3MonthsDate = this._getIntervalDateFromStartDate(this._getIntervalDate("months", 3), 'minutes', 1);

                if (selectedStartDate >= past24Date) {
                    return DateFilterInterval.Past24.toString();
                } else if (selectedStartDate >= pastWeekDate) {
                    return DateFilterInterval.PastWeek.toString();
                } else if (selectedStartDate >= pastMonthDate) {
                    return DateFilterInterval.PastMonth.toString();
                } else if (selectedStartDate >= past3MonthsDate) {
                    return DateFilterInterval.Past3Months.toString();
                } else if (selectedStartDate >= pastYearDate) {
                    return DateFilterInterval.PastYear.toString();
                }
            }
        } else {
            return DateFilterInterval.AnyTime.toString();
        }
    }

    private _updateFromDate(fromDate: Date) {

        this.setState({
            selectedFromDate: fromDate,
            enableApplyDate: true
        });
    }

    private _updateToDate(toDate: Date) {

        this.setState({
            selectedToDate: toDate,
            enableApplyDate: true
        });
    }

    private _onApplyDate = () => {

        this._updateFilter(this.state.selectedFromDate, this.state.selectedToDate, true);
    }

    private _updateFilter(selectedFromDate: Date, selectedToDate: Date, selected: boolean) {

        let updatedValues: IDataFilterValueInfo[] = [];

        let startDate = selectedFromDate ? selectedFromDate.toISOString() : null;
        let endDate = selectedToDate ? selectedToDate.toISOString() : null;

        // Build values
        if (startDate) {
            updatedValues.push({
                name: startDate,
                value: startDate,
                operator: FilterComparisonOperator.Geq,
                selected: selected
            });
        }

        if (endDate) {
            updatedValues.push({
                name: endDate,
                value: endDate,
                operator: FilterComparisonOperator.Lt,
                selected: selected
            });
        }

        this.props.onUpdate(updatedValues);
    }

    private _onFormatDate(date: Date): string {
        return this.props.moment(date).format('LL');
    }

    private _renderMenuList(
        menuListProps: IContextualMenuListProps,
        defaultRender: IRenderFunction<IContextualMenuListProps>
    ) {
        const datePickerStyles = (props: IDatePickerStyleProps) => {
            const customStyles: Partial<IDatePickerStyles> = {
                textField: {
                    selectors: {
                        input: {
                            backgroundColor: this.props.themeVariant.semanticColors.bodyBackground,
                            color: this.props.themeVariant.semanticColors.bodyText,
                            border: "1px solid"
                        },
                        'input::placeholder': {
                            color: this.props.themeVariant.semanticColors.bodyText
                        }
                    }
                },
                root: {
                    padding: "0px 0px 10px 0"
                }
            };

            return customStyles;
        };

        const fromProps: IDatePickerProps = {
            label: "From",
            placeholder: "Select a date...",
            ariaLabel: "Select a date",
            value: this.state.selectedFromDate,
            onSelectDate: this._updateFromDate,
            showGoToToday: true,
            borderless: true,
            styles: datePickerStyles,
            theme: this.props.themeVariant as ITheme,
            strings: strings.General.DatePickerStrings,
            formatDate: this._onFormatDate,
            allowTextInput: true
        };

        const toProps: IDatePickerProps = {
            label: "To",
            placeholder: "Select a date...",
            ariaLabel: "Select a date",
            value: this.state.selectedToDate,
            onSelectDate: this._updateToDate,
            showGoToToday: true,
            borderless: true,
            styles: datePickerStyles,
            theme: this.props.themeVariant as ITheme,
            strings: strings.General.DatePickerStrings,
            formatDate: this._onFormatDate,
            allowTextInput: true
        };

        if (this.state.selectedFromDate) {
            const minDdate = new Date(this.state.selectedFromDate.getTime());
            minDdate.setDate(this.state.selectedFromDate.getDate() + 1);
            toProps.minDate = minDdate;
        }

        if (this.state.selectedToDate) {
            const maxDate = new Date(this.state.selectedToDate.getTime());
            maxDate.setDate(this.state.selectedToDate.getDate() - 1);
            fromProps.maxDate = maxDate;
        }

        return (
            <div>
                {defaultRender(menuListProps)}
                < div style={{ padding: "16px 14px", boxSizing: "border-box" }}>
                    <DatePicker {...fromProps} />
                    <DatePicker {...toProps} />
                    <div style={{ margin: "5px 0" }} >
                        <PrimaryButton text="Apply dates" allowDisabledFocus width={100} style={{ width: "100%" }} disabled={!this.state.enableApplyDate} onClick={this._onApplyDate} />
                    </div>
                </div >
            </div >
        );
    }
}

export class FilterDateTimeWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        const dateHelper = this._serviceScope.consume<DateHelper>(DateHelper.ServiceKey);
        const moment = await dateHelper.moment();

        let props = this.resolveAttributes();
        let renderDateTime: JSX.Element = null;

        if (props.filter) {

            const filter = props.filter as IDataFilterInternal;
            renderDateTime = <FilterDateTimeComponent {...props} moment={moment} filter={filter} onUpdate={((filterValues: IDataFilterValueInfo[]) => {

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
            }).bind(this)} />;
        } else {
            renderDateTime = <MessageBar
                messageBarType={MessageBarType.warning}
                isMultiline={false}>
                {`Component <pnp-date-range> misconfigured. The HTML attribute 'filter' is missing.`}
            </MessageBar>;
        }

        ReactDOM.render(renderDateTime, this);
    }
}