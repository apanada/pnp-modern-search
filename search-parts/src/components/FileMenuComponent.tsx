import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ITheme, CommandBarButton, IButtonStyles, IContextualMenuItem, IContextualMenuItemProps, IImageProps, Image, Callout, Text, FocusZone, PrimaryButton, DefaultButton, Stack, FocusTrapCallout, mergeStyleSets, FontWeights, FocusZoneTabbableElements, Link, Toggle, Spinner, SpinnerSize, TextField, Label } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { UrlHelper } from '../helpers/UrlHelper';
import { ServiceScope } from '@microsoft/sp-core-library';
import { ISharePointSearchService } from '../services/searchService/ISharePointSearchService';
import { SharePointSearchService } from '../services/searchService/SharePointSearchService';
import { isEmpty } from '@microsoft/sp-lodash-subset';

const SUPPORTED_OFFICE_EXTENSIONS: string[] = [
    "doc", "docx", "docm", "dot", "dotm", "dotx",
    "xls", "xlsx", "xlsm",
    "ppt", "pptx", "pptm", "pot", "potm", "pps", "ppsx",
    "vsd", "vsdx", "vss", "vdx", "vsdm", "vsdx"
];

export interface IFileMenuProps {

    /**
    * The file extension
    */
    extension?: string;

    /**
     * Flag indicating if the item is a container (ex: folder)
     */
    isContainer?: string;

    /**
     * The icon size
     */
    size?: string;

    /**
     * The search result item
     */
    resultItem?: any;

    /**
     * The current theme settings
     */
    themeVariant?: IReadonlyTheme;

    /**
     * The current service scope reference
     */
    serviceScope: ServiceScope;
}

export interface IFileMenuState {
    checkAccessCalloutVisible?: boolean;

    userHasAccessToReport?: boolean;

    reportsDocumentSetItemCount?: number;

    isLoading?: boolean;

    requestForAccess?: boolean;
}

const checkAccessStyles = mergeStyleSets({
    callout: {
        width: 320,
        padding: '0px 0px',
    },
    title: {
        marginBottom: 12,
        fontWeight: FontWeights.semilight,
    },
    buttons: {
        display: 'flex',
        justifyContent: 'flex-end',
        marginTop: 10,
    },
    link: {
        display: 'inline-flex',
        marginTop: 5,
    }
});

export class FileMenu extends React.Component<IFileMenuProps, IFileMenuState> {

    /**
     * Image url to use as the icon
     */
    private imageUrl?: string;

    private redirectUrl?: string;

    private originalPath?: string;

    private siteId?: string;

    private siteUrl?: string;

    private uniqueId?: string;

    private parentLink?: string;

    constructor(props: IFileMenuProps) {
        super(props);

        this.state = {
            isLoading: false,
            checkAccessCalloutVisible: false,
            userHasAccessToReport: false,
            reportsDocumentSetItemCount: 0,
            requestForAccess: false
        };

        this._openDocumentInBrowser = this._openDocumentInBrowser.bind(this);
        this._openDocumentInApp = this._openDocumentInApp.bind(this);
        this._downloadDocument = this._downloadDocument.bind(this);
        this._openDocumentParentFolder = this._openDocumentParentFolder.bind(this);
        this._searchThisSite = this._searchThisSite.bind(this);
        this._checkReportAccess = this._checkReportAccess.bind(this);
        this._requestForAccess = this._requestForAccess.bind(this);
        this._toggleCheckAccessCallout = this._toggleCheckAccessCallout.bind(this);
    }

    public componentDidMount(): void {
        if (this.props.resultItem) {
            this.imageUrl = this.props.resultItem["siteLogo"];
            this.redirectUrl = this.props.resultItem["serverRedirectedURL"];
            this.originalPath = this.props.resultItem["originalPath"];
            this.siteUrl = this.props.resultItem["spSiteURL"];
            this.siteId = this.props.resultItem["normSiteID"];
            this.uniqueId = this.props.resultItem["normUniqueID"];
            this.parentLink = this.props.resultItem["parentLink"];
        }
    }

    public render() {

        const styles: Partial<IButtonStyles> = {
            root: {
                width: "24px",
                height: "24px",
                lineHeight: "24px",
                verticalAlign: "middle",
                border: "none",
                color: this.props.themeVariant.palette.themePrimary,
                cursor: "pointer",
                display: "inline-block",
                padding: 0,
                textAlign: "center",
                minWidth: "24px",
            },
            label: {
                fontSize: "20px",
                height: "32px",
                padding: 0,
                textAlign: "center",
                width: "32px"
            },
            menuIcon: {
                display: "none",
                opacity: 0
            }
        };

        const menuItems: IContextualMenuItem[] = [];

        const openInBrowserMenuItem: IContextualMenuItem = {
            key: 'openInBrowser',
            text: 'Open in browser',
            onRenderIcon: (props: IContextualMenuItemProps) => {
                const imageProps: Partial<IImageProps> = {
                    src: this._getOfficeBrandIcons(this.props.extension),
                    styles: properties => ({ root: { color: properties.theme.palette.neutralSecondary, fontSize: "16px", width: "16px" } })
                };

                return (
                    <span>
                        <Image {...imageProps} alt="Open in browser" />
                    </span>
                );
            },
            onClick: this._openDocumentInBrowser
        };

        const openInAppMenuItem: IContextualMenuItem = {
            key: 'openInApp',
            text: 'Open in app',
            onRenderIcon: (props: IContextualMenuItemProps) => {
                const imageProps: Partial<IImageProps> = {
                    src: this._getOfficeBrandIcons(this.props.extension),
                    styles: properties => ({ root: { color: properties.theme.palette.neutralSecondary, fontSize: "16px", width: "16px" } })
                };

                return (
                    <span>
                        <Image {...imageProps} alt="Open in app" />
                    </span>
                );
            },
            onClick: this._openDocumentInApp
        };

        if (SUPPORTED_OFFICE_EXTENSIONS.includes(this.props.extension)) {
            menuItems.push(openInBrowserMenuItem);
            menuItems.push(openInAppMenuItem);
        }

        //If Content Type === <Shell Document Set> OR Content Type === <Shell Report>
        if (this.props.resultItem &&
            !isEmpty(this.props.resultItem["contentType"]) &&
            (this.props.resultItem["contentType"] === "Document Set" || this.props.resultItem["contentType"] === "Shell Report")) {
            menuItems.push({
                key: 'checkReportAccess',
                iconProps: { iconName: 'Signin' },
                text: 'Check access',
                onClick: this._checkReportAccess
            });
        }

        menuItems.push(
            {
                key: 'openFolder',
                iconProps: { iconName: 'FolderHorizontal' },
                text: 'Open folder',
                onClick: this._openDocumentParentFolder
            },
            {
                key: 'download',
                iconProps: { iconName: 'Download' },
                text: 'Download',
                onClick: this._downloadDocument
            },
            {
                key: 'searchThisSite',
                iconProps: { iconName: 'Search' },
                text: 'Search this site',
                onClick: this._searchThisSite
            }
        );

        const uniqueId = Math.floor(Math.random() * 1000) + 1;

        return <div>
            <div>
                <CommandBarButton
                    id={`results-menu-item-${uniqueId}`}
                    text="..."
                    styles={styles}
                    theme={this.props.themeVariant as ITheme}
                    menuProps={{
                        shouldFocusOnMount: true,
                        items: [...menuItems],
                    }} />
            </div>
            <div>
                {this.state.checkAccessCalloutVisible && (
                    <Callout
                        role="checkaccesscallout"
                        ariaLabelledBy="check-access-callout-label"
                        className={checkAccessStyles.callout}
                        gapSpace={0}
                        target={`#results-menu-item-${uniqueId}`}
                        isBeakVisible={true}
                        setInitialFocus
                        onDismiss={this._toggleCheckAccessCallout}
                        styles={{
                            calloutMain: {
                                padding: "20px 24px"
                            }
                        }}
                    >
                        <Text block variant="xLarge" className={checkAccessStyles.title}>
                            Reports: Check Permissions
                        </Text>
                        {
                            this.state.isLoading ?
                                this.state.userHasAccessToReport ?
                                    <div>
                                        <Text block variant="small" styles={{ root: { fontSize: "13px" } }}>
                                            You have access to this report, please click the below link to view the report.
                                        </Text>
                                        <Link target="_blank" onClick={this._openLinkInNewTab.bind(this, `${this.originalPath}/Contents`)} className={checkAccessStyles.link}>
                                            Shell Report
                                        </Link>
                                        <div>
                                            {
                                                <Text block variant="small" styles={{ root: { fontSize: "13px" } }}>
                                                    <Label styles={{ root: { color: this.props.themeVariant.palette.themePrimary } }}>Report Type | {this.state.reportsDocumentSetItemCount > 0 ? "Electronic" : "Physical"}</Label>
                                                </Text>
                                            }
                                        </div>
                                    </div>
                                    :
                                    <div>
                                        <Text block variant="small" styles={{ root: { fontSize: "13px" } }}>
                                            You don't have access to this report.
                                        </Text>
                                        <Toggle label="Would you like to request for your access?" inlineLabel onText="Yes" offText="No" styles={{ label: { fontSize: "13px" }, text: { fontSize: "13px" } }} onChange={this._requestForAccess} />
                                        {
                                            this.state.requestForAccess &&
                                            <TextField label="Request comments:" multiline autoAdjustHeight defaultValue="I'd like access, please." />
                                        }
                                    </div>
                                :
                                <div>
                                    <Spinner size={SpinnerSize.large} />
                                </div>
                        }

                        <FocusZone handleTabKey={FocusZoneTabbableElements.all} isCircularNavigation>
                            <Stack className={checkAccessStyles.buttons} gap={8} horizontal>
                                {
                                    this.state.requestForAccess &&
                                    <PrimaryButton onClick={this._toggleCheckAccessCallout}>Request Access</PrimaryButton>
                                }
                                <DefaultButton onClick={this._toggleCheckAccessCallout}>Cancel</DefaultButton>
                            </Stack>
                        </FocusZone>
                    </Callout>
                )}
            </div>
        </div>;
    }

    private _getOfficeBrandIcons = (extension: string): string | undefined => {
        let brandIcon: string = undefined;

        switch (extension) {
            case "doc":
            case "docx":
            case "docm":
            case "dot":
            case "dotx":
                brandIcon = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/word_32x1.svg";
                break;
            case "xls":
            case "xlsx":
            case "xlsm":
                brandIcon = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_32x1.svg";
                break;
            case "ppt":
            case "pptx":
            case "pptm":
            case "pot":
            case "potm":
            case "pps":
            case "ppsx":
                brandIcon = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/powerpoint_32x1.svg";
                break;
            case "vsd":
            case "vsdx":
            case "vss":
            case "vdx":
            case "vsdm":
            case "vsdx":
                brandIcon = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/visio_32x1.svg";
                break;
            default:
                brandIcon = "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/word_32x1.svg";
                break;
        }

        return brandIcon;
    }

    private _getOfficeClientAppScheme = (extension: string): string | undefined => {
        let clientAppScheme: string = undefined;

        switch (extension) {
            case "doc":
            case "docx":
            case "docm":
            case "dot":
            case "dotx":
                clientAppScheme = "ms-word";
                break;
            case "xls":
            case "xlsx":
            case "xlsm":
                clientAppScheme = "ms-excel";
                break;
            case "ppt":
            case "pptx":
            case "pptm":
            case "pot":
            case "potm":
            case "potm":
            case "potm":
            case "potm":
            case "potm":
            case "ppsx":
                clientAppScheme = "ms-powerpoint";
                break;
            case "vsd":
            case "vsdx":
            case "vss":
            case "vdx":
            case "vsdm":
            case "vsdx":
                clientAppScheme = "ms-visio";
                break;
            default:
                clientAppScheme = "ms-word";
                break;
        }

        return clientAppScheme;
    }

    /**
     * Opens the document in a new tab. The code use window.open
     */
    private _openDocumentInBrowser(ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
        let newTabObject: any = null;
        try {
            if (this.redirectUrl) {
                let documentWebUrl: string = this.redirectUrl;
                newTabObject = window.open(documentWebUrl);
            }
        }
        catch (ex) {
            //optionaly, we can notify the user;
            // cuurently - do nothing
        }
    }

    /**
     * Opens the document in client application
     */
    private _openDocumentInApp(ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
        let newTabObject: any = null;
        try {
            if (this.props.extension && this.originalPath) {
                let clientAppUrl: string = `${this._getOfficeClientAppScheme(this.props.extension)}:ofe|u|${this.originalPath}`;
                newTabObject = window.open(clientAppUrl, "_self");
            }
        }
        catch (ex) {
            //optionaly, we can notify the user;
            // cuurently - do nothing
        }
    }

    /**
     * Downloads the document. The code use window.open
     */
    public _downloadDocument(ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
        let newTabObject: any = null;
        try {
            if (this.siteUrl && this.uniqueId) {
                let documentDownloadUrl: string = `${this.siteUrl}/_layouts/15/download.aspx?UniqueId=${this.uniqueId}`;
                newTabObject = window.open(documentDownloadUrl, "_self");
            }
        }
        catch (ex) {
            //optionaly, we can notify the user;
            // cuurently - do nothing
        }
    }

    /**
     * Opens the document parent folder. The code use window.open
     */
    private _openDocumentParentFolder(ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
        let newTabObject: any = null;
        try {
            if (this.parentLink) {
                let parentFolderURl: string = this.parentLink;
                newTabObject = window.open(parentFolderURl);
            }
        }
        catch (ex) {
            //optionaly, we can notify the user;
            // cuurently - do nothing
        }
    }

    private _searchThisSite(ev?: React.MouseEvent<HTMLElement, MouseEvent> | React.KeyboardEvent<HTMLElement>, item?: IContextualMenuItem): void {
        let newTabObject: any = null;
        try {
            if (this.siteId) {
                let searchThisSiteUrl: string = UrlHelper.addOrReplaceQueryStringParam(window.location.href, "scope", "site");
                searchThisSiteUrl = UrlHelper.addOrReplaceQueryStringParam(searchThisSiteUrl, "sid", this.props.toString());

                newTabObject = window.open(searchThisSiteUrl, "_blank");
            }
        }
        catch (ex) {
            //optionaly, we can notify the user;
            // cuurently - do nothing
        }
    }

    private _toggleCheckAccessCallout() {
        this.setState({
            isLoading: false,
            checkAccessCalloutVisible: false,
            reportsDocumentSetItemCount: 0,
            requestForAccess: false,
            userHasAccessToReport: false
        });
    }

    private _checkReportAccess() {

        if (this.originalPath) {
            this.setState({ checkAccessCalloutVisible: true });

            const sharePointSearchService = this.props.serviceScope.consume<ISharePointSearchService>(SharePointSearchService.ServiceKey);
            const userHasAccess = sharePointSearchService.checkUserAccessToReports(this.originalPath);
            userHasAccess.then(hasAccess => {
                this.setState({
                    isLoading: true,
                    reportsDocumentSetItemCount: hasAccess.ItemCount,
                    userHasAccessToReport: hasAccess.hasAccess
                });
            });
        }
    }

    private _openLinkInNewTab(link: string) {
        let newTabObject: any = null;
        try {
            if (link) {
                newTabObject = window.open(link, "_blank");
            }
        }
        catch (ex) {
            //optionaly, we can notify the user;
            // cuurently - do nothing
        }
    }

    private _requestForAccess(ev: React.MouseEvent<HTMLElement>, checked?: boolean) {
        this.setState({
            requestForAccess: checked
        });
    }
}

export class FileMenuWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        const fileMenu = <FileMenu {...props} serviceScope={this._serviceScope} />;
        ReactDOM.render(fileMenu, this);
    }
}