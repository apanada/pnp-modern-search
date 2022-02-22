import * as React from 'react';
import { IPreviewContainerProps, PreviewType } from './IPreviewContainerProps';
import IPreviewContainerState from './IPreviewContainerState';
import { CommandBarButton, DefaultButton, Dialog, DialogType, IButtonStyles, Icon, IconButton, IIconProps, ILabelStyles, IModalProps, IOverflowSetItemProps, IStackTokens, IStyleSet, Label, Link, OverflowSet, Pivot, PivotItem, PivotLinkFormat, Stack } from 'office-ui-fabric-react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import previewContainerStyles from './PreviewContainer.module.scss';
import { Overlay } from 'office-ui-fabric-react';
import { TestConstants } from '../../common/Constants';
import { split } from 'lodash';

const modalPropsStyles = { main: { maxWidth: 1300 } };
const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
    root: { marginTop: 10 },
};

export default class PreviewContainer extends React.Component<IPreviewContainerProps, IPreviewContainerState> {

    public constructor(props: IPreviewContainerProps) {
        super(props);
        this.state = {
            showDialog: false,
            isLoading: true,
            isFollowed: false
        };

        this._onCloseCallout = this._onCloseCallout.bind(this);
        this._userActions = this._userActions.bind(this);
        this._followOrUnfollowDocument = this._followOrUnfollowDocument.bind(this);
    }

    private _userActions(): JSX.Element {

        const followedIcon: IIconProps = { iconName: 'FavoriteStarFill' };
        const unfollowedIcon: IIconProps = { iconName: 'FavoriteStar' };

        return (
            <OverflowSet
                aria-label="userActions"
                role="menubar"
                items={[
                    {
                        key: 'follow',
                        onRender: () => {
                            return (
                                <div>
                                    <span>
                                        <DefaultButton
                                            toggle
                                            text={this.state.isFollowed ? 'Unfollow' : 'Follow'}
                                            iconProps={this.state.isFollowed ? followedIcon : unfollowedIcon}
                                            onClick={this._followOrUnfollowDocument}
                                            allowDisabledFocus
                                            style={{ border: "none" }}
                                        />
                                    </span>
                                </div>
                            );
                        },
                    }
                ]}
                onRenderOverflowButton={onRenderOverflowButton}
                onRenderItem={onRenderItem}
            />
        );
    }

    public render(): React.ReactElement<IPreviewContainerProps> {
        const { showDialog } = this.state;
        let renderPreview: JSX.Element = null;

        switch (this.props.previewType) {
            case PreviewType.Document:
                renderPreview = <div data-ui-test-id={TestConstants.PreviewCallout} className={`${previewContainerStyles.iframeContainer} ${this.state.isLoading ? previewContainerStyles.hide : ''}`}>
                    <iframe
                        src={this.props.elementUrl} frameBorder="0"
                        allowTransparency
                        onLoad={() => { this.setState({ isLoading: false }); }}
                    >
                    </iframe>
                </div>;
                break;

            default:
                break;
        }

        let renderLoading: JSX.Element = this.state.isLoading ? <Overlay isDarkThemed={false} className={previewContainerStyles.overlay}><Spinner size={SpinnerSize.large} /></Overlay> : null;

        let backgroundImage = this.state.isLoading ? `url('${this.props.previewImageUrl}')` : 'none';

        // Stack tokens definition
        const stackTokens: IStackTokens = {
            childrenGap: 5,
            padding: 10,
        };

        // Dialog props definition
        const dialogContentProps = {
            type: DialogType.largeHeader,
            title: this.props.resultItem["resource"]["fields"]["filename"]
        };
        const modalProps: IModalProps = {
            isBlocking: false,
            topOffsetFixed: false,
            styles: modalPropsStyles,
            dragOptions: undefined,
        };

        const createdDate: string = this._getDate(this.props.resultItem["resource"]["createdDateTime"]);
        const lastModifiedDate: string = this._getDate(this.props.resultItem["resource"]["lastModifiedDateTime"]);
        let author: string = this.props.resultItem["resource"]["fields"]["authorOWSUSER"];
        if (author && split(author, "|").length > 1) {
            author = split(author, "|")[1];
        }

        let authors: string[] = [];
        const metadataAuthors: string = this.props.resultItem["resource"]["fields"]["metadataAuthor"];
        if (metadataAuthors) {
            authors = split(metadataAuthors, "\n\n");
        }

        return (
            <Dialog
                hidden={!showDialog}
                onDismiss={this.props.previewType === PreviewType.Document ? this._onCloseCallout : null}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
                minWidth="1300px"
            >
                <div style={{ display: "flex", justifyContent: "flex-end" }}>
                    {this._userActions()}
                </div>
                <div>
                    <Pivot
                        aria-label="Select an option"
                        linkFormat={PivotLinkFormat.links}
                    >
                        <PivotItem headerText="Document Preview" itemIcon="RedEye">
                            <div className={previewContainerStyles.calloutContentContainer} style={{ backgroundImage: backgroundImage }}>
                                {renderLoading}
                                {renderPreview}
                            </div>
                        </PivotItem>
                        <PivotItem headerText="Metadata" itemIcon="Tag">
                            <div>
                                <Stack horizontal tokens={stackTokens}>
                                    <div className={previewContainerStyles.keyValueList}>
                                        {
                                            this.props.resultItem["resource"]["fields"]["title"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Title:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["resource"]["fields"]["title"] ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            this.props.resultItem["resource"]["fields"]["fileType"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>File Type:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["resource"]["fields"]["fileType"] ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            author &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Created By:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{author ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            createdDate &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Created Date:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{createdDate ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                    </div>
                                    <div className={previewContainerStyles.keyValueList}>
                                        {
                                            this.props.resultItem["resource"]["fields"]["description"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Description:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["resource"]["fields"]["description"] ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            this.props.resultItem["resource"]["fields"]["modifiedBy"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Modified By:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["resource"]["fields"]["modifiedBy"] ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            lastModifiedDate &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Last Modified Time:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{lastModifiedDate ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            this.props.resultItem["resource"]["fields"]["documentLink"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>DocumentLink:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["resource"]["fields"]["filename"] ?? ""}</span>
                                                        <Link href={this.props.resultItem["resource"]["fields"]["documentLink"]} target='_blank' style={{ marginLeft: "8px" }}>
                                                            <Icon iconName="OpenInNewTab" title="Open in new tab" ariaLabel="Open in new tab" />
                                                        </Link>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                    </div>
                                    <div className={previewContainerStyles.keyValueList}>
                                        {
                                            this.props.resultItem["resource"]["fields"]["siteTitle"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Site Title:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["resource"]["fields"]["siteTitle"]}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            this.props.resultItem["resource"]["fields"]["size"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>File Size:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this._formatBytes(this.props.resultItem["resource"]["fields"]["size"] ?? 0).toString()}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            this.props.resultItem["resource"]["fields"]["metadataAuthor"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Authors:</Label>
                                                    </div>
                                                    <div className="keyValueValue" style={{ paddingTop: "5px" }}>
                                                        {
                                                            authors && authors.map((authorItem: string) => (
                                                                <>
                                                                    <span className={previewContainerStyles.pill}>{authorItem}</span>
                                                                </>
                                                            ))
                                                        }
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                    </div>
                                </Stack>
                            </div>
                        </PivotItem>
                    </Pivot>
                </div>
            </Dialog>
        );
    }

    public async componentDidMount() {
        this.setState({
            showDialog: this.props.showPreview,
            isLoading: true
        });

        const originalPath = this.props.resultItem["resource"]["fields"]["originalPath"];
        if (originalPath) {
            const isFollowed = await this.props.sharePointSearchService.isDocumentFollowed(originalPath);
            this.setState({
                isFollowed: isFollowed
            });
        }
    }

    public componentWillReceiveProps(nextProps: IPreviewContainerProps) {
        this.setState({
            showDialog: nextProps.showPreview
        });
    }

    private _onCloseCallout() {
        this.setState({
            showDialog: false
        });
    }

    private _getDate = (date: string): string => {
        try {
            if (date) {
                let itemDate = new Date(date);
                const month = itemDate.toLocaleString('default', { month: 'long' });
                const day = itemDate.getDate();
                const year = itemDate.getFullYear();
                return `${month} ${day}, ${year}`;
            }
            return date;
        } catch (error) {
            return date;
        }
    }

    private _formatBytes = (bytes: number, decimals: number = 2): string => {
        if (bytes === 0) {
            return '0 Bytes';
        }
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }

    private _followOrUnfollowDocument = () => {
        this.setState({
            isFollowed: !this.state.isFollowed,
        }, async () => {
            const originalPath = this.props.resultItem["resource"]["fields"]["originalPath"];
            if (originalPath) {
                if (this.state.isFollowed) {
                    const isFollowed = await this.props.sharePointSearchService.followDocument(originalPath);
                } else {
                    const isFollowed = await this.props.sharePointSearchService.stopFollowingDocument(originalPath);
                }
            }
        });
    }
}

const onRenderItem = (item: IOverflowSetItemProps): JSX.Element => {
    if (item.onRender) {
        return item.onRender(item);
    }
    return (
        <CommandBarButton
            role="menuitem"
            iconProps={{ iconName: item.icon }}
            menuProps={item.subMenuProps}
            text={item.name}
        />
    );
};

const onRenderOverflowButton = (overflowItems: any[] | undefined): JSX.Element => {
    const buttonStyles: Partial<IButtonStyles> = {
        root: {
            minWidth: 0,
            padding: '0 4px',
            alignSelf: 'stretch',
            height: 'auto',
        },
    };
    return (
        <CommandBarButton
            ariaLabel="More items"
            role="menuitem"
            styles={buttonStyles}
            menuIconProps={{ iconName: 'More' }}
            menuProps={{ items: overflowItems! }}
        />
    );
};