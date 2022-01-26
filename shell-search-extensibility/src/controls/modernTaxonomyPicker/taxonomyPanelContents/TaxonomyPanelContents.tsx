import * as React from 'react';
import styles from './TaxonomyPanelContents.module.scss';
import { Checkbox, ICheckboxStyleProps, ICheckboxStyles } from 'office-ui-fabric-react/lib/Checkbox';
import { ChoiceGroup, IChoiceGroupOption, IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles } from 'office-ui-fabric-react/lib/ChoiceGroup';
import {
  GroupedList,
  GroupHeader,
  IGroup,
  IGroupFooterProps,
  IGroupHeaderProps,
  IGroupHeaderStyleProps,
  IGroupHeaderStyles,
  IGroupRenderProps,
  IGroupShowAllProps
} from 'office-ui-fabric-react/lib/GroupedList';
import { IBasePickerStyleProps, IBasePickerStyles, IPickerItemProps, ISuggestionItemProps } from 'office-ui-fabric-react/lib/Pickers';
import {
  ILabelStyleProps,
  ILabelStyles,
  Label
} from 'office-ui-fabric-react/lib/Label';
import {
  ILinkStyleProps,
  ILinkStyles,
  Link
} from 'office-ui-fabric-react/lib/Link';
import { IListProps } from 'office-ui-fabric-react/lib/List';
import { IRenderFunction, IStyleFunctionOrObject, Selection, SelectionMode } from 'office-ui-fabric-react/lib/Utilities';
import { ISpinnerStyleProps, ISpinnerStyles, Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { SelectionZone } from 'office-ui-fabric-react/lib/Selection';
import {
  ITermInfo,
  ITermSetInfo,
  ITermStoreInfo
} from '@pnp/sp/taxonomy';
import { Guid } from '@microsoft/sp-core-library';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { css } from '@uifabric/utilities/lib/css';
import * as strings from 'ControlStrings';
import { useForceUpdate } from '@uifabric/react-hooks';
import { ModernTermPicker } from '../modernTermPicker/ModernTermPicker';
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { IModernTermPickerProps } from '../modernTermPicker/ModernTermPicker.types';
import { Optional } from '../ModernTaxonomyPicker';
import { IButtonStyles, IconButton, IIconProps } from 'office-ui-fabric-react';

export interface ITaxonomyPanelContentsProps {
  context: BaseComponentContext;
  allowMultipleSelections?: boolean;
  termSetId: Guid;
  pageSize: number;
  selectedPanelOptions: ITermInfo[];
  setSelectedPanelOptions: React.Dispatch<React.SetStateAction<ITermInfo[]>>;
  onResolveSuggestions: (filter: string, selectedItems?: ITermInfo[]) => ITermInfo[] | PromiseLike<ITermInfo[]>;
  onLoadMoreData: (termSetId: Guid, parentTermId?: Guid, skiptoken?: string, hideDeprecatedTerms?: boolean, pageSize?: number) => Promise<{ value: ITermInfo[], skiptoken: string }>;
  anchorTermInfo: ITermInfo;
  termSetInfo: ITermSetInfo;
  termStoreInfo: ITermStoreInfo;
  placeHolder: string;
  onRenderSuggestionsItem?: (props: ITermInfo, itemProps: ISuggestionItemProps<ITermInfo>) => JSX.Element;
  onRenderItem?: (props: IPickerItemProps<ITermInfo>) => JSX.Element;
  getTextFromItem: (item: ITermInfo, currentValue?: string) => string;
  languageTag: string;
  themeVariant?: IReadonlyTheme;
  termPickerProps?: Optional<IModernTermPickerProps, 'onResolveSuggestions'>;
}

export function TaxonomyPanelContents(props: ITaxonomyPanelContentsProps): React.ReactElement<ITaxonomyPanelContentsProps> {
  const [groupsLoading, setGroupsLoading] = React.useState<string[]>([]);
  const [groups, setGroups] = React.useState<IGroup[]>([]);
  const [terms, setTerms] = React.useState<ITermInfo[]>(props.selectedPanelOptions?.length > 0 ? [...props.selectedPanelOptions] : []);

  const forceUpdate = useForceUpdate();

  const selection = React.useMemo(() => {
    const s = new Selection({
      onSelectionChanged: () => {
        props.setSelectedPanelOptions((prevOptions) => [...selection.getSelection()]);
        forceUpdate();
      }, getKey: (term: any) => term.id
    });
    s.setItems(terms);
    for (const selectedOption of props.selectedPanelOptions) {
      if (s.canSelectItem) {
        s.setKeySelected(selectedOption.id.toString(), true, true);
      }
    }
    return s;
  }, [terms]);

  React.useEffect(() => {
    let termRootName = "";
    if (props.anchorTermInfo) {
      let anchorTermNames = props.anchorTermInfo.labels.filter((name) => name.languageTag === props.languageTag && name.isDefault);
      if (anchorTermNames.length === 0) {
        anchorTermNames = props.anchorTermInfo.labels.filter((name) => name.languageTag === props.termStoreInfo.defaultLanguageTag && name.isDefault);
      }
      termRootName = anchorTermNames[0].name;
    }
    else {
      let termSetNames = props.termSetInfo.localizedNames.filter((name) => name.languageTag === props.languageTag);
      if (termSetNames.length === 0) {
        termSetNames = props.termSetInfo.localizedNames.filter((name) => name.languageTag === props.termStoreInfo.defaultLanguageTag);
      }
      termRootName = termSetNames[0].name;
    }
    const rootGroup: IGroup = {
      name: termRootName,
      key: props.anchorTermInfo ? props.anchorTermInfo.id : props.termSetInfo.id,
      startIndex: -1,
      count: 50,
      level: 0,
      isCollapsed: false,
      data: { skiptoken: '' },
      hasMoreData: (props.anchorTermInfo ? props.anchorTermInfo.childrenCount : props.termSetInfo.childrenCount) > 0
    };
    setGroups([rootGroup]);
    setGroupsLoading((prevGroupsLoading) => [...prevGroupsLoading, props.termSetInfo.id]);
    if (props.termSetInfo.childrenCount > 0) {
      props.onLoadMoreData(props.termSetId, props.anchorTermInfo ? Guid.parse(props.anchorTermInfo.id) : Guid.empty, '', true)
        .then((loadedTerms) => {
          const grps: IGroup[] = loadedTerms.value.map(term => {
            let termNames = term.labels.filter((termLabel) => (termLabel.languageTag === props.languageTag && termLabel.isDefault === true));
            if (termNames.length === 0) {
              termNames = term.labels.filter((termLabel) => (termLabel.languageTag === props.termStoreInfo.defaultLanguageTag && termLabel.isDefault === true));
            }
            const g: IGroup = {
              name: termNames[0]?.name,
              key: term.id,
              startIndex: -1,
              count: 50,
              level: 1,
              isCollapsed: true,
              data: { skiptoken: '', term: term },
              hasMoreData: term.childrenCount > 0,
            };
            if (g.hasMoreData) {
              g.children = [];
            }
            return g;
          });
          setTerms((prevTerms) => {
            const nonExistingTerms = loadedTerms.value.filter((term) => prevTerms.every((prevTerm) => prevTerm.id !== term.id));
            return [...prevTerms, ...nonExistingTerms];
          });
          rootGroup.children = grps;
          rootGroup.data.skiptoken = loadedTerms.skiptoken;
          rootGroup.hasMoreData = loadedTerms.skiptoken !== '';
          setGroupsLoading((prevGroupsLoading) => prevGroupsLoading.filter((value) => value !== props.termSetId.toString()));
          setGroups([rootGroup]);
        });
    }
  }, []);

  const onToggleCollapse = (group: IGroup): void => {
    if (group.isCollapsed === true) {
      setGroups((prevGroups) => {
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

        return newGroupsState;
      });

      if (group.children && group.children.length === 0) {
        setGroupsLoading((prevGroupsLoading) => [...prevGroupsLoading, group.key]);
        group.data.isLoading = true;

        props.onLoadMoreData(props.termSetId, Guid.parse(group.key), '', true)
          .then((loadedTerms) => {
            const grps: IGroup[] = loadedTerms.value.map(term => {
              let termNames = term.labels.filter((termLabel) => (termLabel.languageTag === props.languageTag && termLabel.isDefault === true));
              if (termNames.length === 0) {
                termNames = term.labels.filter((termLabel) => (termLabel.languageTag === props.termStoreInfo.defaultLanguageTag && termLabel.isDefault === true));
              }
              const g: IGroup = {
                name: termNames[0]?.name,
                key: term.id,
                startIndex: -1,
                count: 50,
                level: group.level + 1,
                isCollapsed: true,
                data: { skiptoken: '', term: term },
                hasMoreData: term.childrenCount > 0,
              };
              if (g.hasMoreData) {
                g.children = [];
              }
              return g;
            });

            setTerms((prevTerms) => {
              const nonExistingTerms = loadedTerms.value.filter((term) => prevTerms.every((prevTerm) => prevTerm.id !== term.id));
              return [...prevTerms, ...nonExistingTerms];
            });

            group.children = grps;
            group.data.skiptoken = loadedTerms.skiptoken;
            group.hasMoreData = loadedTerms.skiptoken !== '';
            setGroupsLoading((prevGroupsLoading) => prevGroupsLoading.filter((value) => value !== group.key));
          });
      }
    }
    else {
      setGroups((prevGroups) => {
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

        return newGroupsState;
      });

    }
  };

  const onRenderTitle = (groupHeaderProps: IGroupHeaderProps) => {
    const isChildSelected = (children: IGroup[]): boolean => {
      let aChildIsSelected = children && children.some((child) => selection.isKeySelected(child.key) || isChildSelected(child.children));
      return aChildIsSelected;
    };

    const childIsSelected = isChildSelected(groupHeaderProps.group.children);

    if (groupHeaderProps.group.level === 0) {
      const labelStyles: IStyleFunctionOrObject<ILabelStyleProps, ILabelStyles> = { root: { fontWeight: childIsSelected ? "bold" : "normal" } };
      const addTermButtonStyles: IButtonStyles = { rootHovered: { backgroundColor: 'transparent' }, rootPressed: { backgroundColor: 'transparent' }, icon: { color: "#000000", cursor: "auto" }, root: { position: "absolute"} };
      return (
        <>
          <IconButton styles={addTermButtonStyles} iconProps={{ iconName: 'Package' } as IIconProps} />
          <Label styles={labelStyles}>{groupHeaderProps.group.name}</Label>
        </>
      );
    }

    const isDisabled = groupHeaderProps.group.data.term.isAvailableForTagging.filter((t) => t.setId === props.termSetId.toString())[0].isAvailable === false;
    const isSelected = selection.isKeySelected(groupHeaderProps.group.key);

    const selectionProps = {
      "data-selection-index": selection.getItems().findIndex((term) => term.id === groupHeaderProps.group.key)
    };

    if (props.allowMultipleSelections) {
      if (isDisabled) {
        selectionProps["data-selection-disabled"] = true;
      }
      else {
        selectionProps["data-selection-toggle"] = true;
      }

      const selectedStyles: IStyleFunctionOrObject<ICheckboxStyleProps, ICheckboxStyles> = { root: { pointerEvents: 'none' } };
      if (isSelected || childIsSelected) {
        selectedStyles.label = { fontWeight: 'bold' };
      }
      else {
        selectedStyles.label = { fontWeight: 'normal' };
      }

      return (
        <div {...selectionProps}>
          <Checkbox
            key={groupHeaderProps.group.key}
            label={groupHeaderProps.group.name}
            checked={isSelected}
            styles={selectedStyles}
            disabled={isDisabled}
            onRenderLabel={(p) => <span className={css(!isDisabled && styles.checkbox, isDisabled && styles.disabledCheckbox, isSelected && styles.selectedCheckbox)} title={p.title}>
              {p.label}
            </span>}
          />
        </div>
      );
    }
    else {
      const selectedStyle: IStyleFunctionOrObject<IChoiceGroupOptionStyleProps, IChoiceGroupOptionStyles> = isSelected || childIsSelected ? { root: { marginTop: 0 }, choiceFieldWrapper: { fontWeight: 'bold', } } : { root: { marginTop: 0 }, choiceFieldWrapper: { fontWeight: 'normal' } };
      const options: IChoiceGroupOption[] = [{
        key: groupHeaderProps.group.key,
        text: groupHeaderProps.group.name,
        styles: selectedStyle,
        onRenderLabel: (p) =>
          <span id={p.labelId} className={css(!isDisabled && styles.choiceOption, isDisabled && styles.disabledChoiceOption, isSelected && styles.selectedChoiceOption)}>
            {p.text}
          </span>
      }];

      if (isDisabled) {
        selectionProps["data-selection-disabled"] = true;
      }
      else {
        selectionProps["data-selection-select"] = true;
      }

      return (
        <div {...selectionProps}>
          <ChoiceGroup
            options={options}
            selectedKey={selection.getSelection()[0]?.id}
            disabled={isDisabled}
          />
        </div>
      );
    }
  };

  const onRenderHeader = (headerProps: IGroupHeaderProps): JSX.Element => {
    const groupHeaderStyles: IStyleFunctionOrObject<IGroupHeaderStyleProps, IGroupHeaderStyles> = {
      expand: { height: 42, visibility: !headerProps.group.children || headerProps.group.level === 0 ? "hidden" : "visible", fontSize: 14 },
      expandIsCollapsed: { visibility: !headerProps.group.children || headerProps.group.level === 0 ? "hidden" : "visible", fontSize: 14 },
      check: { display: 'none' },
      headerCount: { display: 'none' },
      groupHeaderContainer: { height: 36, paddingTop: 3, paddingBottom: 3, paddingLeft: 3, paddingRight: 3, alignItems: 'center', },
      root: { height: 42 },
    };

    return (
      <GroupHeader
        {...headerProps}
        styles={groupHeaderStyles}
        onRenderTitle={onRenderTitle}
        onToggleCollapse={onToggleCollapse}
        indentWidth={20}
        expandButtonProps={{ style: { color: props.themeVariant?.semanticColors.bodyText } }}
      />
    );
  };

  const onRenderFooter = (footerProps: IGroupFooterProps): JSX.Element => {
    if ((footerProps.group.hasMoreData || footerProps.group.children && footerProps.group.children.length === 0) && !footerProps.group.isCollapsed) {

      if (groupsLoading.some(value => value === footerProps.group.key)) {
        const spinnerStyles: IStyleFunctionOrObject<ISpinnerStyleProps, ISpinnerStyles> = { circle: { verticalAlign: 'middle' } };
        return (
          <div className={styles.spinnerContainer}>
            <Spinner styles={spinnerStyles} />
          </div>
        );
      }
      const linkStyles: IStyleFunctionOrObject<ILinkStyleProps, ILinkStyles> = { root: { fontSize: '14px', paddingLeft: (footerProps.groupLevel + 1) * 20 + 62 } };
      return (
        <div className={styles.loadMoreContainer}>
          <Link onClick={() => {
            setGroupsLoading((prevGroupsLoading) => [...prevGroupsLoading, footerProps.group.key]);
            props.onLoadMoreData(props.termSetId, footerProps.group.key === props.termSetId.toString() ? Guid.empty : Guid.parse(footerProps.group.key), footerProps.group.data.skiptoken, true)
              .then((loadedTerms) => {
                const grps: IGroup[] = loadedTerms.value.map(term => {
                  let termNames = term.labels.filter((termLabel) => (termLabel.languageTag === props.languageTag && termLabel.isDefault === true));
                  if (termNames.length === 0) {
                    termNames = term.labels.filter((termLabel) => (termLabel.languageTag === props.termStoreInfo.defaultLanguageTag && termLabel.isDefault === true));
                  }
                  const g: IGroup = {
                    name: termNames[0]?.name,
                    key: term.id,
                    startIndex: -1,
                    count: 50,
                    level: footerProps.group.level + 1,
                    isCollapsed: true,
                    data: { skiptoken: '', term: term },
                    hasMoreData: term.childrenCount > 0,
                  };
                  if (g.hasMoreData) {
                    g.children = [];
                  }
                  return g;
                });
                setTerms((prevTerms) => {
                  const nonExistingTerms = loadedTerms.value.filter((term) => prevTerms.every((prevTerm) => prevTerm.id !== term.id));
                  return [...prevTerms, ...nonExistingTerms];
                });
                footerProps.group.children = [...footerProps.group.children, ...grps];
                footerProps.group.data.skiptoken = loadedTerms.skiptoken;
                footerProps.group.hasMoreData = loadedTerms.skiptoken !== '';
                setGroupsLoading((prevGroupsLoading) => prevGroupsLoading.filter((value) => value !== footerProps.group.key));
              });
          }}
            styles={linkStyles}>
            {strings.ModernTaxonomyPickerLoadMoreText}
          </Link>
        </div>
      );
    }
    return null;
  };

  const onRenderShowAll: IRenderFunction<IGroupShowAllProps> = () => {
    return null;
  };

  const groupProps: IGroupRenderProps = {
    onRenderFooter: onRenderFooter,
    onRenderHeader: onRenderHeader,
    showEmptyGroups: true,
    onRenderShowAll: onRenderShowAll,
  };

  const onPickerChange = (items?: ITermInfo[]): void => {
    const itemsToAdd = items.filter((item) => terms.every((term) => term.id !== item.id));
    setTerms((prevTerms) => [...prevTerms, ...itemsToAdd]);
    selection.setItems([...selection.getItems(), ...itemsToAdd], true);
    for (const item of items) {
      if (selection.canSelectItem(item)) {
        selection.setKeySelected(item.id.toString(), true, false);
      }
    }
  };

  const termPickerStyles: IStyleFunctionOrObject<IBasePickerStyleProps, IBasePickerStyles> = { root: { paddingTop: 4, paddingBottom: 4, paddingRight: 4, minheight: 34 }, input: { minheight: 34 }, text: { minheight: 34, borderStyle: 'none', borderWidth: '0px' } };

  return (
    <div className={styles.taxonomyPanelContents}>
      <div className={styles.taxonomyTreeSelector}>
        <div>
          <ModernTermPicker
            {...props.termPickerProps}
            removeButtonAriaLabel={strings.ModernTaxonomyPickerRemoveButtonText}
            onResolveSuggestions={props.termPickerProps?.onResolveSuggestions ?? props.onResolveSuggestions}
            itemLimit={props.allowMultipleSelections ? undefined : 1}
            selectedItems={props.selectedPanelOptions}
            styles={props.termPickerProps?.styles ?? termPickerStyles}
            onChange={onPickerChange}
            getTextFromItem={props.getTextFromItem}
            pickerSuggestionsProps={props.termPickerProps?.pickerSuggestionsProps ?? { noResultsFoundText: strings.ModernTaxonomyPickerNoResultsFound }}
            inputProps={props.termPickerProps?.inputProps ?? {
              'aria-label': props.placeHolder || strings.ModernTaxonomyPickerDefaultPlaceHolder,
              placeholder: props.placeHolder || strings.ModernTaxonomyPickerDefaultPlaceHolder
            }}
            onRenderSuggestionsItem={props.termPickerProps?.onRenderSuggestionsItem ?? props.onRenderSuggestionsItem}
            onRenderItem={props.onRenderItem ?? props.onRenderItem}
            themeVariant={props.themeVariant}
          />
        </div>
      </div>
      <div>
        <SelectionZone
          selectionMode={props.allowMultipleSelections ? SelectionMode.multiple : SelectionMode.single}
          selection={selection}
        >
          <GroupedList
            items={[]}
            onRenderCell={null}
            groups={groups}
            groupProps={groupProps}
            onShouldVirtualize={(p: IListProps<any>) => false}
            selectionMode={props.allowMultipleSelections ? SelectionMode.multiple : SelectionMode.single}
          />
        </SelectionZone>
      </div>
    </div>
  );
}

