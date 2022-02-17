import { ServiceScope } from "@microsoft/sp-core-library";
import { ITermGroupInfo, ITermSetInfo, ITermInfo, ITermStoreInfo } from "@pnp/sp/taxonomy";
import { ChoiceGroup, DefaultButton, DetailsRow, Dialog, DialogFooter, DialogType, FocusZone, GroupedList, GroupHeader, IButtonStyles, IChoiceGroupOption, IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles, IColumn, Icon, IconButton, IGroup, IGroupFooterProps, IGroupHeaderProps, IGroupHeaderStyleProps, IGroupHeaderStyles, IGroupRenderProps, IGroupShowAllProps, IIconProps, IIconStyles, ILabelStyleProps, ILabelStyles, IListProps, Label, PrimaryButton, SelectionZone } from "office-ui-fabric-react";
import * as React from "react";
import { ITaxonomyService } from "../services/taxonomyService/ITaxonomyService";
import { TaxonomyService } from "../services/taxonomyService/TaxonomyService";
import { IRenderFunction, IStyleFunctionOrObject, Selection, SelectionMode } from 'office-ui-fabric-react/lib/Utilities';
import { PageContext } from "@microsoft/sp-page-context";
import { isEmpty } from "@microsoft/sp-lodash-subset";

export interface ITermSetPickerResult {

    /**
     * Selected termset id.
     */
    termSetId: string;
}

export interface ITermSetPickerProps {

    /**
     * The current service scope reference
     */
    serviceScope: ServiceScope;

    /**
     * Specifies the label of the termset picker dialog
     */
    dialogLabel?: string;

    /**
   * Handler when the termset has been selected
   */
    onSave?: (termSetPickerResult: ITermSetPickerResult) => void;

    /**
     * Handler when file picker has been cancelled.
     */
    onCancel?: () => void;

    /**
     * Specifies if the termset picker dialog is open by default or not
     */
    isDialogOpen?: boolean;
}

export interface ITermSetPickerState {
    /**
     * Specifies the term store details
     */
    termStoreInfo?: ITermStoreInfo;

    /**
     * Specifies the current language tag
     */
    languageTag?: string;


    /**
     * Specifies whether to show or hide termset picker dialog
     */
    showDialog: boolean;

    /**
     * Specifies the top level term groups
     */
    termGroups?: ITermGroupInfo[];

    /**
     * Specifies the top level termsets
     */
    termSets?: ITermSetInfo[];

    /**
     * Specifies the terms selected
     */
    terms?: any[];

    /**
     * Specifies the groups displayed in the termset picker
     */
    groups?: IGroup[];

    /**
     * Specifies the selected termset option from termset picker
     */
    selectedOption?: IGroup;
}

export class TermSetPicker extends React.Component<ITermSetPickerProps, ITermSetPickerState> {

    private _pageContext: PageContext;
    private _taxonomyService: ITaxonomyService;
    private _selection: Selection;

    constructor(props: ITermSetPickerProps) {
        super(props);

        this.state = {
            showDialog: this.props.isDialogOpen || false,
            termGroups: [],
            termSets: [],
            terms: [],
            groups: []
        };

        this._pageContext = props.serviceScope.consume<PageContext>(PageContext.serviceKey);
        this._taxonomyService = this.props.serviceScope.consume<ITaxonomyService>(TaxonomyService.ServiceKey);

        this._selection = new Selection({
            onSelectionChanged: this._onItemsSelectionChanged,
            getKey: (term: any) => term.id
        });
    }

    public async componentDidMount() {

        if (this._taxonomyService) {
            const termStoreInfo = await this._taxonomyService.getTermStoreInfo();
            const languageTag = this._pageContext.cultureInfo.currentUICultureName !== '' ?
                this._pageContext.cultureInfo.currentUICultureName :
                termStoreInfo.defaultLanguageTag;

            const termGroups: ITermGroupInfo[] = await this._taxonomyService.getTermGroups();

            const groups: IGroup[] = termGroups.map((group, index) => {
                const g: IGroup = {
                    name: group.name,
                    key: group.id,
                    startIndex: -1,
                    count: 50,
                    level: 1,
                    isCollapsed: true,
                    data: { group: group },
                    children: []
                };
                return g;
            });

            this.setState({
                termStoreInfo: termStoreInfo,
                languageTag: languageTag,
                termGroups: termGroups,
                groups: groups
            });
        }
    }

    /**
     * componentWillReceiveProps lifecycle hook
     *
     * @param nextProps
     */
    public componentWillReceiveProps(nextProps: ITermSetPickerProps): void {
        if (nextProps.isDialogOpen || nextProps.isDialogOpen !== this.props.isDialogOpen) {
            this.setState({
                showDialog: nextProps.isDialogOpen
            });
        }
    }

    public render(): JSX.Element {

        const { showDialog } = this.state;

        const modelProps = {
            isBlocking: true,
            styles: { main: { maxWidth: "500px !important" } },
        };
        const dialogContentProps = {
            type: DialogType.largeHeader,
            title: this.props.dialogLabel,
            subText: 'Select term set you want to display as filters for the current filter.',
        };

        return (
            <>
                <div>
                    <Icon iconName="Tag" title="Select Termset" aria-label="Tag" onClick={this._handleOpenDialog} />
                </div>
                <div>
                    <Dialog
                        hidden={!showDialog}
                        onDismiss={this._toggleHideDialog}
                        dialogContentProps={dialogContentProps}
                        modalProps={modelProps}
                    >
                        <div>
                            <FocusZone>
                                <SelectionZone
                                    selection={this._selection}
                                    selectionMode={SelectionMode.single}
                                >
                                    <GroupedList
                                        items={[]}
                                        onRenderCell={null}
                                        selection={this._selection}
                                        selectionMode={SelectionMode.single}
                                        groups={this.state.groups}
                                        groupProps={{
                                            onRenderHeader: this.onRenderHeader.bind(this),
                                            onRenderShowAll: this.onRenderShowAll.bind(this),
                                            showEmptyGroups: true
                                        }}
                                        onShouldVirtualize={(p: IListProps<any>) => false}
                                    />
                                </SelectionZone>
                            </FocusZone>
                        </div>
                        <DialogFooter>
                            <PrimaryButton onClick={this._onSave} text="Save" />
                            <DefaultButton onClick={this._toggleHideDialog} text="Cancel" />
                        </DialogFooter>
                    </Dialog>
                </div>
            </>
        );
    }

    private _onItemsSelectionChanged = () => {
        this.forceUpdate();
    }

    private onRenderHeader(headerProps: IGroupHeaderProps): JSX.Element {
        const groupHeaderStyles: IStyleFunctionOrObject<IGroupHeaderStyleProps, IGroupHeaderStyles> = {
            expand: { height: 42, visibility: !headerProps.group.children || headerProps.group.level === 0 ? "hidden" : "visible", fontSize: 14 },
            expandIsCollapsed: { visibility: !headerProps.group.children || headerProps.group.level === 0 ? "hidden" : "visible", fontSize: 14 },
            check: { display: 'none' },
            headerCount: { display: 'none' },
            groupHeaderContainer: { height: 36, paddingTop: 3, paddingBottom: 3, paddingLeft: 3, paddingRight: 3, alignItems: 'center', },
            root: { height: 42 },
        };

        const onToggleSelectGroup = () => {
            headerProps.onToggleCollapse(headerProps.group);
            if (headerProps.group.level > 1) {
                if (this._selection.canSelectItem(headerProps.group)) {
                    this._selection.setItems([...this.state.terms], true);
                    this._selection.setKeySelected(headerProps.group.key.toString(), true, false);
                    this.setState({
                        selectedOption: headerProps.group
                    });
                }
            }
        };

        const onToggleCollapse = (group: IGroup): void => {

            if (group.isCollapsed === true) {
                const prevGroups = this.state.groups;
                const recurseGroups = (currentGroup: IGroup) => {
                    if (currentGroup.key === group.key) {
                        currentGroup.isCollapsed = false;
                    }
                    if (currentGroup.children?.length > 0) {
                        for (const child of currentGroup.children) {
                            recurseGroups(child);
                        }
                    }
                };

                let newGroupsState: IGroup[] = [];
                for (const prevGroup of prevGroups) {
                    recurseGroups(prevGroup);
                    newGroupsState.push(prevGroup);
                }

                this.setState({
                    groups: newGroupsState
                });

                if (group.children && group.children.length === 0) {
                    this._taxonomyService.getTermSets(group.key).then(termSets => {
                        const grps: IGroup[] = termSets.map((termSet, index) => {
                            let termSetNames = termSet.localizedNames.filter((localizedName) => (localizedName.languageTag === this.state.languageTag));
                            if (termSetNames.length === 0) {
                                termSetNames = termSet.localizedNames.filter((localizedName) => (localizedName.languageTag === this.state.termStoreInfo.defaultLanguageTag));
                            }
                            const g: IGroup = {
                                name: termSetNames[0]?.name,
                                key: termSet.id,
                                startIndex: -1,
                                count: 50,
                                level: group.level + 1,
                                isCollapsed: true,
                                data: { termSet: termSet },
                                hasMoreData: false
                            };
                            return g;
                        });

                        group.children = grps;

                        const prevTerms = this.state.terms;
                        const nonExistingTerms = termSets.filter((term) => prevTerms.every((prevTerm) => prevTerm.id !== term.id));
                        this.setState({ terms: [...prevTerms, ...nonExistingTerms] }, () => {
                            const itemsToAdd = this.state.terms.map(term => {
                                let termSetNames = term.localizedNames.filter((localizedName) => (localizedName.languageTag === this.state.languageTag));
                                if (termSetNames.length === 0) {
                                    termSetNames = term.localizedNames.filter((localizedName) => (localizedName.languageTag === this.state.termStoreInfo.defaultLanguageTag));
                                }
                                return {
                                    key: termSetNames[0]?.name,
                                    name: termSetNames[0]?.name,
                                    id: term.id
                                };
                            });

                            this._selection.setItems([...itemsToAdd], false);
                        });
                    });
                }
            } else {
                const prevGroups = this.state.groups;
                const recurseGroups = (currentGroup: IGroup) => {
                    if (currentGroup.key === group.key) {
                        currentGroup.isCollapsed = true;
                    }
                    if (currentGroup.children?.length > 0) {
                        for (const child of currentGroup.children) {
                            recurseGroups(child);
                        }
                    }
                };

                let newGroupsState: IGroup[] = [];
                for (const prevGroup of prevGroups) {
                    recurseGroups(prevGroup);
                    newGroupsState.push(prevGroup);
                }

                this.setState({
                    groups: newGroupsState
                });
            }
        };

        const onRenderTitle = (groupHeaderProps: IGroupHeaderProps) => {

            const isChildSelected = (children: IGroup[]): boolean => {
                let aChildIsSelected = children && children.some((child) => this._selection.isKeySelected(child.key) || isChildSelected(child.children));
                return aChildIsSelected;
            };

            const childIsSelected = isChildSelected(groupHeaderProps.group.children);

            if (groupHeaderProps.group.level === 1) {

                const labelStyles: IStyleFunctionOrObject<ILabelStyleProps, ILabelStyles> = { root: { fontWeight: childIsSelected ? "bold" : "normal" } };
                const addTermButtonStyles: IIconStyles = { root: { margin: "2px 8px 2px 0px", color: "#000000", cursor: "auto", alignItems: "center", verticalAlign: "middle" } };
                return (
                    <>
                        <Label styles={labelStyles}>
                            <Icon styles={addTermButtonStyles} iconName='FabricFolder' />
                            {groupHeaderProps.group.name}
                        </Label>
                    </>
                );
            }

            const isSelected = this._selection.isKeySelected(groupHeaderProps.group.key);

            const selectionProps = {
                "data-_selection-index": this._selection.getItems().findIndex((term: any) => term.id === groupHeaderProps.group.key)
            };

            const selectedStyle: IStyleFunctionOrObject<IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles> = isSelected || childIsSelected ? { root: { marginTop: 0 }, choiceFieldWrapper: { fontWeight: 'bold', } } : { root: { marginTop: 0 }, choiceFieldWrapper: { fontWeight: 'normal' } };
            const options: IChoiceGroupOption[] = [{
                key: groupHeaderProps.group.key,
                text: groupHeaderProps.group.name,
                styles: selectedStyle,
                onRenderLabel: (p) =>
                    <span id={p.labelId}
                        style={{
                            color: "[theme: bodyText, default: #323130]",
                            display: "inline-block",
                            paddingInlineStart: "26px"
                        }}>
                        {p.text}
                    </span>
            }];

            return (
                <div {...selectionProps}>
                    <ChoiceGroup
                        options={options}
                        selectedKey={this._selection.getSelection() && this._selection.getSelection().length > 0 ? this._selection.getSelection()[0]["id"] : null}
                    />
                </div>
            );
        };

        return (
            <GroupHeader
                {...headerProps}
                styles={groupHeaderStyles}
                onRenderTitle={onRenderTitle.bind(this)}
                onToggleSelectGroup={onToggleSelectGroup.bind(this)}
                onToggleCollapse={onToggleCollapse.bind(this)}
                indentWidth={20}
            />
        );
    }

    private onRenderShowAll: IRenderFunction<IGroupShowAllProps> = () => {
        return null;
    }

    private _onSave = () => {
        if (!isEmpty(this.state.selectedOption)) {
            this.props.onSave({
                termSetId: this.state.selectedOption.key
            });
            this._toggleHideDialog();
        }
    }

    /**
     * Open the dialog
     */
    private _handleOpenDialog = () => {
        this.setState({
            showDialog: true
        });
    }

    private _toggleHideDialog = () => {
        this.setState({
            showDialog: false
        });
    }
}