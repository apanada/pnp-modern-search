import * as React from 'react';
import styles from './SearchVerticalsContainer.module.scss';
import { ISearchVerticalsContainerProps } from './ISearchVerticalsContainerProps';
import { Pivot, PivotItem, IPivotItemProps, Icon, GlobalSettings, IChangeDescription, ITheme, ActionButton, Dialog, DialogFooter, PrimaryButton, DefaultButton, DialogType, DialogContent, IChoiceGroupOption, ChoiceGroup, TextField, Text, Link, DirectionalHint } from '@fluentui/react';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { PageOpenBehavior } from '../../../helpers/UrlHelper';
import { ISearchVerticalsContainerState } from './ISearchVerticalsContainerState';
import { BuiltinTokenNames } from '../../../services/tokenService/TokenService';
import { TooltipHost } from '@microsoft/office-ui-fabric-react-bundle';
import { isEmpty } from '@microsoft/sp-lodash-subset';

const feedbackModelProps = {
  isBlocking: false,
  topOffsetFixed: false,
  styles: { main: { maxWidth: 450 } }
};

const feedbackDialogContentProps = {
  type: DialogType.largeHeader,
  title: 'Tell us what you think about search.',
  styles: { title: { fontSize: "18px" } },
  showCloseButton: true
};

export default class SearchVerticalsContainer extends React.Component<ISearchVerticalsContainerProps, ISearchVerticalsContainerState> {

  public constructor(props) {
    super(props);

    this.state = {
      selectedKey: undefined,
      hideFeedbackDialog: true,
      selectedFeedbackEmoji: undefined,
      selectedFeedbackOption: undefined
    };

    // Listen to inputQueryText value change on the page
    GlobalSettings.addChangeListener((changeDescription: IChangeDescription) => {
      if (changeDescription.key === BuiltinTokenNames.inputQueryText) {
        this.props.tokenService.setTokenValue(BuiltinTokenNames.inputQueryText, GlobalSettings.getValue(BuiltinTokenNames.inputQueryText));
      }
    });

    this.onVerticalSelected = this.onVerticalSelected.bind(this);
    this._toggleHideFeedbackDialog = this._toggleHideFeedbackDialog.bind(this);
    this._onFeedbackOptionChange = this._onFeedbackOptionChange.bind(this);
    this._dynamicFeedbackOptions = this._dynamicFeedbackOptions.bind(this);
  }

  public render(): React.ReactElement<ISearchVerticalsContainerProps> {

    const { hideFeedbackDialog, selectedFeedbackEmoji, selectedFeedbackOption } = this.state;
    let renderTitle: JSX.Element = null;

    // Web Part title
    renderTitle = <WebPartTitle
      displayMode={this.props.webPartTitleProps.displayMode}
      title={this.props.webPartTitleProps.title}
      updateProperty={this.props.webPartTitleProps.updateProperty}
      className={this.props.webPartTitleProps.className}
    />;

    const renderPivotItems = this.props.verticals.map(vertical => {

      let pivotItemProps: IPivotItemProps = {};
      let renderLinkIcon: JSX.Element = null;

      if (vertical.iconName && vertical.iconName.trim() !== "") {
        pivotItemProps.itemIcon = vertical.iconName;
      }

      if (vertical.showLinkIcon) {
        renderLinkIcon = vertical.openBehavior === PageOpenBehavior.NewTab ?
          <Icon styles={{ root: { fontSize: 10, paddingLeft: 3 } }} iconName='NavigateExternalInline'></Icon> :
          <Icon styles={{ root: { fontSize: 10, paddingLeft: 3 } }} iconName='Link'></Icon>;
      }

      return <PivotItem
        headerText={vertical.tabName}
        itemKey={vertical.key}
        onRenderItemLink={(props, defaultRender) => {

          if (vertical.isLink) {
            return <div className={styles.isLink}>
              {defaultRender(props)}
              {renderLinkIcon}
            </div>;
          } else {
            return defaultRender(props);
          }
        }}
        {...pivotItemProps}>
      </PivotItem>;
    });

    return <>
      {renderTitle}
      <div className={styles.searchVerticals}>
        <Pivot
          className={styles.dataVerticals}
          onLinkClick={this.onVerticalSelected}
          selectedKey={this.state.selectedKey}
          theme={this.props.themeVariant as ITheme}
          overflowBehavior='menu'
          overflowAriaLabel="more items"
        >
          {renderPivotItems}
        </Pivot>
        <div className={styles.toolTipStyles}>
          <div className={styles.feedbackButtonStyles}>
            <TooltipHost content="Provide search feedback" id="feedback-tooltip" calloutProps={{ gapSpace: 0 }} styles={{ root: { display: 'inline-block' } }}>
              <ActionButton
                iconProps={{ iconName: 'feedback' }}
                allowDisabledFocus
                onClick={this._toggleHideFeedbackDialog}
                aria-describedby="feedback-tooltip"
                styles={{
                  label: {
                    fontSize: "14px",
                    fontWeight: 600,
                    color: "#006cbe"
                  }
                }}
              >
                Feedback
              </ActionButton>
            </TooltipHost>
            <Dialog
              hidden={hideFeedbackDialog}
              minWidth={412}
              onDismiss={this._toggleHideFeedbackDialog}
              modalProps={feedbackModelProps}
              dialogContentProps={feedbackDialogContentProps}>
              <div>
                <div className={styles.feedbackEmojiStyles}>
                  <TooltipHost content="Very satisfied" id="verysatisfied-tooltip" calloutProps={{ gapSpace: 0 }} directionalHint={DirectionalHint.bottomCenter} styles={{ root: { display: 'inline-block' } }}>
                    <span className={selectedFeedbackEmoji === 'Blush' ? styles.activeFeedbackEmoji : null} aria-describedby="verysatisfied-tooltip" onClick={this._onFeedbackEmojiClick.bind(this, 'Blush')}>&#128512;</span>
                  </TooltipHost>
                  <TooltipHost content="Satisfied" id="satisfied-tooltip" calloutProps={{ gapSpace: 0 }} directionalHint={DirectionalHint.bottomCenter} styles={{ root: { display: 'inline-block' } }}>
                    <span className={selectedFeedbackEmoji === 'Okay' ? styles.activeFeedbackEmoji : null} aria-describedby="satisfied-tooltip" onClick={this._onFeedbackEmojiClick.bind(this, 'Okay')}>&#128578;</span>
                  </TooltipHost>
                  <TooltipHost content="Neutral" id="neutral-tooltip" calloutProps={{ gapSpace: 0 }} directionalHint={DirectionalHint.bottomCenter} styles={{ root: { display: 'inline-block' } }}>
                    <span className={selectedFeedbackEmoji === 'Expressionless' ? styles.activeFeedbackEmoji : null} aria-describedby="neutral-tooltip" onClick={this._onFeedbackEmojiClick.bind(this, 'Expressionless')}>&#128529;</span>
                  </TooltipHost>
                  <TooltipHost content="Dissatisfied" id="dissatisfied-tooltip" calloutProps={{ gapSpace: 0 }} directionalHint={DirectionalHint.bottomCenter} styles={{ root: { display: 'inline-block' } }}>
                    <span className={selectedFeedbackEmoji === 'Disappointed' ? styles.activeFeedbackEmoji : null} aria-describedby="dissatisfied-tooltip" onClick={this._onFeedbackEmojiClick.bind(this, 'Disappointed')}>&#128542;</span>
                  </TooltipHost>
                  <TooltipHost content="Confounded" id="confounded-tooltip" calloutProps={{ gapSpace: 0 }} directionalHint={DirectionalHint.bottomCenter} styles={{ root: { display: 'inline-block' } }}>
                    <span className={selectedFeedbackEmoji === 'Confounded' ? styles.activeFeedbackEmoji : null} aria-describedby="confounded-tooltip" onClick={this._onFeedbackEmojiClick.bind(this, 'Confounded')}>&#128577;</span>
                  </TooltipHost>
                  <TooltipHost content="Astonished" id="astonished-tooltip" calloutProps={{ gapSpace: 0 }} directionalHint={DirectionalHint.bottomCenter} styles={{ root: { display: 'inline-block' } }}>
                    <span className={selectedFeedbackEmoji === 'Astonished' ? styles.activeFeedbackEmoji : null} aria-describedby="astonished-tooltip" onClick={this._onFeedbackEmojiClick.bind(this, 'Astonished')}>&#128562;</span>
                  </TooltipHost>
                  <TooltipHost content="Very dissatisfied" id="verydissatisfied-tooltip" calloutProps={{ gapSpace: 0 }} directionalHint={DirectionalHint.bottomCenter} styles={{ root: { display: 'inline-block' } }}>
                    <span className={selectedFeedbackEmoji === 'Angry' ? styles.activeFeedbackEmoji : null} aria-describedby="verydissatisfied-tooltip" onClick={this._onFeedbackEmojiClick.bind(this, 'Angry')}>&#128544;</span>
                  </TooltipHost>
                </div>
                {
                  selectedFeedbackEmoji &&
                  <div>
                    <ChoiceGroup selectedKey={selectedFeedbackOption} options={this._dynamicFeedbackOptions()} onChange={this._onFeedbackOptionChange} />
                    <TextField label="Leave additional comments or suggestions" multiline rows={3} placeholder="Remember not to include personal info like phone numbers." />
                    <p className={styles.consentContainer}>
                      <span className={styles.consentMessage}>
                        <Text variant='smallPlus'>
                          By pressing submit, your feedback will be used to improve modern search experience. Your IT admin will be able to view this data.
                        </Text>
                      </span>
                    </p>
                  </div>
                }
              </div>
              <DialogFooter>
                <PrimaryButton onClick={this._toggleHideFeedbackDialog} text="Submit" disabled={!isEmpty(selectedFeedbackEmoji) && !isEmpty(selectedFeedbackOption) ? false : true} />
                <DefaultButton onClick={this._toggleHideFeedbackDialog} text="Cancel" />
              </DialogFooter>
            </Dialog>
          </div>
        </div>
      </div >
    </>;
  }

  private _toggleHideFeedbackDialog = () => {
    this.setState({
      hideFeedbackDialog: !this.state.hideFeedbackDialog,
      selectedFeedbackEmoji: undefined,
      selectedFeedbackOption: undefined
    });
  }

  private _onFeedbackEmojiClick = (feedbackEmoji: string) => {
    this.setState({
      selectedFeedbackEmoji: feedbackEmoji
    });
  }

  private _onFeedbackOptionChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption) {
    this.setState({
      selectedFeedbackOption: option.key
    })
  }

  /**
   * Method to return dynamic feedback options
   * @returns Returns the feedback options
   */
  private _dynamicFeedbackOptions = () => {

    // User Feedback Feature
    const satisfiedFeedbackOptions: IChoiceGroupOption[] = [
      { key: 'helpful', text: "It's helpful" },
      { key: 'other', text: "Other" }
    ];

    const neutralFeedbackOptions: IChoiceGroupOption[] = [
      { key: 'incorrect', text: "This is incorrect or irrelevant" },
      { key: 'other', text: "Other" }
    ];

    const dissatisfiedFeedbackOptions: IChoiceGroupOption[] = [
      { key: 'notFound', text: "I didn't find what I was looking for" },
      { key: 'delayInLoad', text: "The result took too long to load" },
      { key: 'other', text: "Other" }
    ];

    const veryDissatisfiedFeedbackOptions: IChoiceGroupOption[] = [
      { key: 'foundBug', text: "I found a bug or error message" },
      { key: 'inappropriateContent', text: "The content was inappropriate" },
      { key: 'other', text: "Other" }
    ];

    const { selectedFeedbackEmoji } = this.state;
    switch (selectedFeedbackEmoji) {
      case 'Blush':
      case 'Okay':
        return satisfiedFeedbackOptions;
      case 'Expressionless':
        return neutralFeedbackOptions;
      case 'Disappointed':
      case 'Confounded':
        return dissatisfiedFeedbackOptions;
      case 'Astonished':
      case 'Angry':
        return veryDissatisfiedFeedbackOptions;
      default:
        return satisfiedFeedbackOptions;
    }
  }

  public onVerticalSelected(item: PivotItem): void {

    const verticalIdx = this.props.verticals.map(vertical => vertical.key).indexOf(item.props.itemKey);

    if (verticalIdx !== -1) {

      const vertical = this.props.verticals[verticalIdx];
      if (vertical.isLink) {
        // Send the query to the new page
        const behavior = vertical.openBehavior === PageOpenBehavior.NewTab ? '_blank' : '_self';
        this.props.tokenService.resolveTokens(vertical.linkUrl).then((resolvedUrl: string) => {
          resolvedUrl = resolvedUrl.replace(/\{searchTerms\}|\{SearchBoxQuery\}/gi, GlobalSettings.getValue(BuiltinTokenNames.inputQueryText));
          window.open(resolvedUrl, behavior);
        });

      } else {

        this.setState({
          selectedKey: item.props.itemKey
        });

        this.props.onVerticalSelected(item.props.itemKey);
      }
    }
  }

  public componentDidMount() {

    let defaultSelectedKey = undefined;

    if (this.props.verticals.length > 0) {
      if (this.props.defaultSelectedKey) {
        defaultSelectedKey = this.props.defaultSelectedKey;
      } else {
        // By default, we select the first one
        defaultSelectedKey = this.props.verticals[0].key;
      }
    }

    this.setState({
      selectedKey: defaultSelectedKey
    });

    // Return the default selected key
    this.props.onVerticalSelected(defaultSelectedKey);
  }
}
