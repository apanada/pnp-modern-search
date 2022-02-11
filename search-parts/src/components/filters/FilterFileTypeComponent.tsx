import * as React from 'react';
import { BaseWebComponent, IDataFilterValueInfo, ExtensibilityConstants, IDataFilterInfo, FilterConditionOperator, IDataFilterInternal, IDataFilterValueInternal } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { IContextualMenuListProps, IRenderFunction, SearchBox, DefaultButton, IChoiceGroupOption, ChoiceGroup, Text, Checkbox, ITheme, MessageBar, MessageBarType, IContextualMenuItem, PrimaryButton, ActionButton, ContextualMenuItemType, Icon } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import styles from './FilterFileTypeComponent.module.scss';
import { FilterSearchBox } from './FilterSearchBoxComponent';
import { FilterMulti, FilterMultiEventCallback } from './FilterMultiComponent';
import * as strings from 'CommonStrings';
import { getFileTypeIconProps } from '@uifabric/file-type-icons';

export interface IFilterFileTypeComponentProps {

    /**
     * The current selected filters. Because we can select values outside of values retrieved from results, we need this information to display the default date picker values correctly after the user selection
     */
    filter: IDataFilterInternal;

    /**
     * The Web Part instance ID from where the filter component belongs
     */
    instanceId?: string;

    /**
     * The current theme settings
     */
    themeVariant?: IReadonlyTheme;

    /**
     * Handler when a filter value is selected
     */
    onChecked: (filterName: string, filterValue: IDataFilterValueInfo) => void;

    /**
     * Callback handlers for filter multi events
     */
    onApply: FilterMultiEventCallback;
}

export interface IFilterFileTypeComponentState {
    filterValues: IDataFilterValueInfo[];
}

export class FilterFileTypeComponent extends React.Component<IFilterFileTypeComponentProps, IFilterFileTypeComponentState> {

    public constructor(props: IFilterFileTypeComponentProps) {
        super(props);

        this.setState({
            filterValues: []
        });

        this._clearFilters = this._clearFilters.bind(this);
    }

    public render() {

        let renderContexualMenu: JSX.Element = <div>
            <DefaultButton
                id="fileTypeContextualMenu"
                text={this.props.filter.displayName}
                menuProps={{
                    onRenderMenuList: this._renderMenuList.bind(this),
                    shouldFocusOnMount: true,
                    items: [{ key: 'divider_1', itemType: ContextualMenuItemType.Divider }]
                }}
                styles={{ root: { border: "0px" }, label: { fontSize: "13px" } }}
            />
        </div>

        return <div className={styles.filterFileType}>
            {renderContexualMenu}
        </div>;
    }

    /**
     * Applies all selected filter values for the current filter
     */
    private _applyFilters() {
        this.props.onApply();
    }

    /**
     * Clears all selected filters for the current refiner
     */
    private _clearFilters() {
        // Determine the removed item according to current selected tags
        const removedItems = this.props.filter.values.filter((selectedValue: IDataFilterValueInternal) => {
            return selectedValue.selected;
        });

        if (removedItems.length > 0) {
            const filterValues: IDataFilterValueInfo[] = removedItems.map(item => {
                item.selected = false;
                return item;
            });
            this.setState({
                filterValues: filterValues
            });
        }
    }

    private _renderMenuList(
        menuListProps: IContextualMenuListProps,
        defaultRender: IRenderFunction<IContextualMenuListProps>
    ) {

        let renderInputs: JSX.Element[] = [];

        let filterMultiSelct: JSX.Element = null;
        if (this.props.filter.isMulti) {
            filterMultiSelct = <div className={styles.filterMultiSelct}>
                <ActionButton
                    theme={this.props.themeVariant as ITheme}
                    disabled={this.props.filter.canClear ? false : true}
                    iconProps={{ iconName: 'Cancel' }}
                    styles={{ root: { color: "#006cbe" } }}
                    onClick={this._clearFilters}>
                    {strings.Filters.DeselectAllFiltersButtonLabel}
                </ActionButton>
                <PrimaryButton
                    className={styles.applyBtn}
                    width={100}
                    style={{ width: "100%" }}
                    disabled={this.props.filter.canApply ? false : true}
                    theme={this.props.themeVariant as ITheme}
                    onClick={this._applyFilters}>
                    {strings.Filters.ApplyAllFiltersButtonLabel}
                </PrimaryButton>
            </div>;
        }

        this.props.filter.values.forEach(filter => {
            let filterValue: IDataFilterValueInfo = {
                name: filter.name,
                value: filter.value,
                selected: filter.selected
            };

            let renderInput: JSX.Element = null;

            if (this.props.filter.isMulti) {
                renderInput = <Checkbox
                    styles={{
                        root: {
                            padding: "5px 10px",
                        },
                        label: {
                            width: '100%'
                        },
                        text: {
                            color: filter.count && filter.count === 0 ? this.props.themeVariant.semanticColors.disabledText : this.props.themeVariant.semanticColors.bodyText
                        }
                    }}
                    theme={this.props.themeVariant as ITheme}
                    defaultChecked={filter.selected}
                    disabled={filter.disabled}
                    title={filterValue.name}
                    label={filterValue.name}
                    onChange={(ev, checked: boolean) => {
                        ev && ev.preventDefault();
                        filterValue.selected = checked;
                        this.props.onChecked(this.props.filter.filterName, filterValue);
                    }}
                    onRenderLabel={(props, defaultRender) => {
                        return (
                            <>
                                <span className={styles.filterValue}>
                                    <Icon
                                        {...getFileTypeIconProps({ extension: props.label, size: 16, imageFileType: 'png' })}
                                        styles={{ root: { paddingRight: "10px" } }} />
                                    <Text block nowrap title={props.label}>{props.label}</Text>
                                </span>
                            </>
                        );
                    }}
                />;
            } else {
                renderInput = <ChoiceGroup
                    styles={{
                        root: {
                            position: 'relative',
                            display: 'flex',
                            paddingRight: 10,
                            paddingLeft: 10,
                            paddingBottom: 7,
                            paddingTop: 7,
                            selectors: {
                                '.ms-ChoiceField': {
                                    marginTop: 0
                                }
                            }
                        }
                    }}
                    key={this.props.filter.filterName}
                    options={[
                        {
                            key: filterValue.value,
                            text: filterValue.name,
                            disabled: filter.disabled,

                            checked: filter.selected
                        }
                    ]}
                    onChange={(ev?: React.FormEvent<HTMLInputElement>, option?: IChoiceGroupOption) => {
                        ev && ev.preventDefault();
                        filterValue.selected = ev.currentTarget.checked;
                        this.props.onChecked(this.props.filter.filterName, filterValue);
                    }}
                />;
            }

            renderInputs.push(renderInput);
        });

        return (
            <>
                <div>
                    <div className={styles.filterOption}>
                        <FilterSearchBox
                            filter={this.props.filter}
                            instanceId={this.props.instanceId}
                            themeVariant={this.props.themeVariant}
                            onFilterValueUpdated={this.props.onChecked}
                        />
                    </div>
                    <div className={styles.filterValuesList}>
                        {renderInputs}
                    </div>
                    <div>
                        {defaultRender(menuListProps)}
                    </div>
                    <div>
                        {filterMultiSelct}
                    </div>
                </div>
            </>
        );
    }
}

export class FilterFileTypeWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        let renderCheckBox: JSX.Element = null;

        if (props.filter) {

            const filter = props.filter as IDataFilterInternal;
            renderCheckBox = <FilterFileTypeComponent {...props} filter={filter}
                onChecked={((filterName: string, filterValue: IDataFilterValueInfo) => {
                    // Bubble event through the DOM
                    this.dispatchEvent(new CustomEvent(ExtensibilityConstants.EVENT_FILTER_UPDATED, {
                        detail: {
                            filterName: filterName,
                            filterValues: [filterValue],
                            instanceId: props.instanceId
                        } as IDataFilterInfo,
                        bubbles: true,
                        cancelable: true
                    }));
                }).bind(this)}
                onApply={(() => {
                    // Bubble event through the DOM
                    this.dispatchEvent(new CustomEvent(ExtensibilityConstants.EVENT_FILTER_APPLY_ALL, {
                        detail: {
                            filterName: props.filterName,
                            instanceId: props.instanceId
                        },
                        bubbles: true,
                        cancelable: true
                    }));
                }).bind(this)}
            />;
        } else {
            renderCheckBox = <MessageBar
                messageBarType={MessageBarType.warning}
                isMultiline={false}>
                {`Component <pnp-date-range> misconfigured. The HTML attribute 'filter' is missing.`}
            </MessageBar>;
        }

        ReactDOM.render(renderCheckBox, this);
    }
}