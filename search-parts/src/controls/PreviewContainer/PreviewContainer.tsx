import * as React from 'react';
import { IPreviewContainerProps, PreviewType } from './IPreviewContainerProps';
import IPreviewContainerState from './IPreviewContainerState';
import { ChoiceGroup, ContextualMenu, DefaultButton, DefaultPalette, Dialog, DialogFooter, DialogType, Icon, IconButton, ILabelStyles, IModalProps, IStackItemStyles, IStackStyles, IStackTokens, IStyleSet, Label, Link, Pivot, PivotItem, PrimaryButton, Stack } from 'office-ui-fabric-react';
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
            isLoading: true
        };

        this._onCloseCallout = this._onCloseCallout.bind(this);
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
            title: this.props.resultItem["Title"]
        };
        const modalProps: IModalProps = {
            isBlocking: false,
            topOffsetFixed: false,
            styles: modalPropsStyles,
            dragOptions: undefined,
        };

        const createdDate: string = this._getDate(this.props.resultItem["Created"]);
        const lastModifiedDate: string = this._getDate(this.props.resultItem["LastModifiedTime"]);
        let author: string = this.props.resultItem["AuthorOWSUSER"];
        if (author && split(author, "|").length > 1) {
            author = split(author, "|")[1];
        }

        let authors: string[] = [];
        const metadataAuthors: string = this.props.resultItem["MetadataAuthor"];
        if (metadataAuthors) {
            authors = split(this.props.resultItem["MetadataAuthor"], "\n\n");
        }

        return (
            <Dialog
                hidden={!showDialog}
                onDismiss={this.props.previewType === PreviewType.Document ? this._onCloseCallout : null}
                dialogContentProps={dialogContentProps}
                modalProps={modalProps}
                minWidth="1300px"
            >
                <div>
                    <Pivot aria-label="Select an option">
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
                                            this.props.resultItem["Title"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Title:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["Title"] ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            this.props.resultItem["FileType"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>File Type:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["FileType"] ?? ""}</span>
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
                                            this.props.resultItem["Description"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Description:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["Description"] ?? ""}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            this.props.resultItem["ModifiedBy"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Modified By:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["ModifiedBy"] ?? ""}</span>
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
                                            this.props.resultItem["DocumentLink"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>DocumentLink:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this.props.resultItem["Filename"] ?? ""}</span>
                                                        <Link href={this.props.resultItem["DocumentLink"]} target='_blank' style={{ marginLeft: "8px" }}>
                                                            <Icon iconName="OpenInNewTab" title="Open in new tab" ariaLabel="Open in new tab" />
                                                        </Link>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                    </div>
                                    <div className={previewContainerStyles.keyValueList}>
                                        {
                                            this.props.resultItem["Size"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>File Size:</Label>
                                                    </div>
                                                    <div className="keyValueValue">
                                                        <span>{this._formatBytes(this.props.resultItem["Size"] ?? 0).toString()}</span>
                                                    </div>
                                                </div>
                                            </div>
                                        }
                                        {
                                            this.props.resultItem["MetadataAuthor"] &&
                                            <div className="keyValueWrapper">
                                                <div>
                                                    <div className="keyValueKey">
                                                        <Label styles={labelStyles}>Authors:</Label>
                                                    </div>
                                                    <div className="keyValueValue" style={{ paddingTop: "5px" }}>
                                                        {
                                                            authors && authors.map((author: string) => (
                                                                <>
                                                                    <span className={previewContainerStyles.pill}>{author}</span>
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

    public componentDidMount() {
        this.setState({
            showDialog: this.props.showPreview,
            isLoading: true
        });
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
                return `${month} ${day}, ${year}`
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
}
