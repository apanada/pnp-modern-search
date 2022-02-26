import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ITheme, CommandBarButton, IButtonStyles, IContextualMenuItem, IContextualMenuItemProps, IImageProps, Image, Callout, Text, FocusZone, PrimaryButton, DefaultButton, Stack, FocusTrapCallout, mergeStyleSets, FontWeights, FocusZoneTabbableElements, Link, Toggle, Spinner, SpinnerSize, TextField, Label, Icon, MessageBar, MessageBarType, ProgressIndicator, format, FontIcon } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { UrlHelper } from '../helpers/UrlHelper';
import { ServiceScope } from '@microsoft/sp-core-library';
import { ISharePointSearchService } from '../services/searchService/ISharePointSearchService';
import { SharePointSearchService } from '../services/searchService/SharePointSearchService';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { IAccessRequestResults, IAccessRequestResultsType, RequestAccessStatus } from '../models/common/IAccessRequest';

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

    checkAccessIsLoading?: boolean;

    requestForAccess?: boolean;

    accessRequestSubmissionInProgress?: boolean;

    accessRequestLogged?: boolean;

    accessRequestResponse?: { message: string | undefined, responseType: IAccessRequestResultsType };

    scanRequestSubmissionInProgress?: boolean;

    scanRequestLogged?: boolean;
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

const ACCESS_REQUEST_VALIDATE_FAILURE_MESSAGE = "Unable to validate your access on this report. Please retry after some time.";
const ACCESS_REQUEST_VALIDATE_SUCCESS_MESSAGE = "You have already requested for this report, Requst status : {0}";
const ACCESS_REQUEST_SUCCESS_MESSAGE = "Your access request has been logged successfully. you will receive the notification once your access request has been proccessed.";
const ACCESS_REQUEST_FAILURE_MESSAGE = "Unable to raise your access request. Please retry after some time."

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
            checkAccessIsLoading: false,
            checkAccessCalloutVisible: false,
            userHasAccessToReport: false,
            reportsDocumentSetItemCount: 0,
            requestForAccess: false,
            accessRequestSubmissionInProgress: false,
            accessRequestLogged: false,
            accessRequestResponse: undefined,
            scanRequestSubmissionInProgress: false,
            scanRequestLogged: false
        };

        this._openDocumentInBrowser = this._openDocumentInBrowser.bind(this);
        this._openDocumentInApp = this._openDocumentInApp.bind(this);
        this._downloadDocument = this._downloadDocument.bind(this);
        this._openDocumentParentFolder = this._openDocumentParentFolder.bind(this);
        this._searchThisSite = this._searchThisSite.bind(this);
        this._checkReportAccess = this._checkReportAccess.bind(this);
        this._toggleRequestForAccess = this._toggleRequestForAccess.bind(this);
        this._toggleCheckAccessCallout = this._toggleCheckAccessCallout.bind(this);
        this._requestReportAccess = this._requestReportAccess.bind(this);
        this._requestReportScan = this._requestReportScan.bind(this);
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
        } else {
            menuItems.push({
                key: 'download',
                iconProps: { iconName: 'Download' },
                text: 'Download',
                onClick: this._downloadDocument
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
                            !this.state.checkAccessIsLoading ?
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
                                            {
                                                this.state.reportsDocumentSetItemCount === 0 ?
                                                    this.state.scanRequestLogged ?
                                                        <div>
                                                            <MessageBar
                                                                messageBarType={this.state.accessRequestResponse.responseType === IAccessRequestResultsType.Success ? MessageBarType.success : MessageBarType.error}
                                                                isMultiline={true}
                                                            >
                                                                {this.state.accessRequestResponse.message}
                                                            </MessageBar>
                                                        </div>
                                                        :
                                                        <TextField label="Requester comments:" multiline autoAdjustHeight defaultValue="Please scan and provide access." />
                                                    : null
                                            }
                                        </div>
                                    </div>
                                    :
                                    this.state.accessRequestLogged ?
                                        <div>
                                            <MessageBar
                                                messageBarType={this.state.accessRequestResponse.responseType === IAccessRequestResultsType.Success ? MessageBarType.success : MessageBarType.error}
                                                isMultiline={true}
                                            >
                                                {this.state.accessRequestResponse.message}
                                            </MessageBar>
                                        </div>
                                        :
                                        <div>
                                            <Text block variant="small" styles={{ root: { fontSize: "13px" } }}>
                                                You don't have access to this report.
                                            </Text>
                                            <Toggle label="Would you like to request for your access?" inlineLabel onText="Yes" offText="No" styles={{ label: { fontSize: "13px" }, text: { fontSize: "13px" } }} onChange={this._toggleRequestForAccess} />
                                            {
                                                this.state.requestForAccess &&
                                                <TextField label="Requester comments:" multiline autoAdjustHeight defaultValue="I'd like access, please." />
                                            }
                                        </div>
                                :
                                <div>
                                    <Spinner size={SpinnerSize.large} label="Verifying your access on this report..." />
                                </div>
                        }
                        {
                            this.state.accessRequestSubmissionInProgress &&
                            <div>
                                <ProgressIndicator label="Submitting access request..." />
                            </div>
                        }
                        {
                            this.state.scanRequestSubmissionInProgress &&
                            <div>
                                <ProgressIndicator label="Submitting scan request..." />
                            </div>
                        }
                        <FocusZone handleTabKey={FocusZoneTabbableElements.all} isCircularNavigation>
                            <Stack className={checkAccessStyles.buttons} gap={8} horizontal>
                                {
                                    this.state.requestForAccess && !this.state.accessRequestLogged &&
                                    <PrimaryButton onClick={this._requestReportAccess}>Request Access</PrimaryButton>
                                }
                                {
                                    this.state.reportsDocumentSetItemCount === 0 && !this.state.scanRequestLogged &&
                                    <PrimaryButton onClick={this._requestReportScan}>Request Scan</PrimaryButton>
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
                searchThisSiteUrl = UrlHelper.addOrReplaceQueryStringParam(searchThisSiteUrl, "sid", this.siteId.toString());

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
            checkAccessIsLoading: false,
            checkAccessCalloutVisible: false,
            reportsDocumentSetItemCount: 0,
            requestForAccess: false,
            userHasAccessToReport: false,
            accessRequestLogged: false,
            accessRequestResponse: undefined,
            accessRequestSubmissionInProgress: false,
            scanRequestLogged: false,
            scanRequestSubmissionInProgress: false
        });
    }

    private _checkReportAccess() {

        if (this.originalPath) {
            this.setState({ checkAccessCalloutVisible: true, checkAccessIsLoading: true });

            const sharePointSearchService = this.props.serviceScope.consume<ISharePointSearchService>(SharePointSearchService.ServiceKey);
            const userHasAccess = sharePointSearchService.checkUserAccessToReports(this.siteUrl, this.originalPath);
            userHasAccess.then(hasAccess => {

                if (hasAccess.hasAccess && this.uniqueId) {
                    const element = (
                        <>
                            <FontIcon aria-label="UnlockSolid" iconName="UnlockSolid" />
                            <span style={{ color: "green" }}>Access permitted</span>
                        </>
                    );
                    ReactDOM.render(element, document.getElementById(this.uniqueId));
                }

                this.setState({
                    checkAccessIsLoading: false,
                    reportsDocumentSetItemCount: hasAccess.ItemCount,
                    userHasAccessToReport: false //hasAccess.hasAccess
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

    private _toggleRequestForAccess(ev: React.MouseEvent<HTMLElement>, checked?: boolean) {
        this.setState({
            requestForAccess: checked,
            accessRequestLogged: false,
            accessRequestResponse: undefined,
            accessRequestSubmissionInProgress: false
        });
    }

    private _requestReportAccess(ev?: any) {

        if (this.originalPath) {
            this.setState({ accessRequestSubmissionInProgress: true });
            var reportNumber = this.originalPath.split('/')[this.originalPath.split('/').length - 1];  // Name of Document Set

            const sharePointSearchService = this.props.serviceScope.consume<ISharePointSearchService>(SharePointSearchService.ServiceKey);
            sharePointSearchService.validateAccessRequest("AccessRequest", reportNumber).then((validateAccessRequestResults: IAccessRequestResults | string) => {
                if (!isEmpty(validateAccessRequestResults)) {
                    if (typeof validateAccessRequestResults === "string") {
                        this.setState({
                            accessRequestLogged: true,
                            accessRequestSubmissionInProgress: false,
                            accessRequestResponse: {
                                message: ACCESS_REQUEST_VALIDATE_FAILURE_MESSAGE,
                                responseType: IAccessRequestResultsType.Failure
                            }
                        });
                    } else {

                        if (this.uniqueId) {
                            const element = (
                                <>
                                    <FontIcon aria-label="CompletedSolid" iconName="CompletedSolid" />
                                    <span style={{ color: "green" }}>Access Request Status : {validateAccessRequestResults.RequestAccessStatus}</span>
                                </>
                            );
                            ReactDOM.render(element, document.getElementById(this.uniqueId));
                        }

                        this.setState({
                            accessRequestLogged: true,
                            accessRequestSubmissionInProgress: false,
                            accessRequestResponse: {
                                message: format(ACCESS_REQUEST_VALIDATE_SUCCESS_MESSAGE, validateAccessRequestResults.RequestAccessStatus),
                                responseType: IAccessRequestResultsType.Success
                            }
                        });
                    }
                } else {
                    sharePointSearchService.submitAccessRequest("AccessRequest", {}).then((accessRequestResults: IAccessRequestResults | string) => {
                        if (isEmpty(accessRequestResults) || (!isEmpty(accessRequestResults) && typeof accessRequestResults === "string")) {
                            this.setState({
                                accessRequestLogged: true,
                                accessRequestSubmissionInProgress: false,
                                accessRequestResponse: {
                                    message: ACCESS_REQUEST_FAILURE_MESSAGE,
                                    responseType: IAccessRequestResultsType.Failure
                                }
                            });
                        } else {

                            if (this.uniqueId) {
                                const element = (
                                    <>
                                        <FontIcon aria-label="CheckMark" iconName="CheckMark" />
                                        <span style={{ color: "green" }}>Access Request Status : {RequestAccessStatus.New}</span>
                                    </>
                                );
                                ReactDOM.render(element, document.getElementById(this.uniqueId));
                            }

                            this.setState({
                                accessRequestLogged: true,
                                accessRequestSubmissionInProgress: false,
                                accessRequestResponse: {
                                    message: ACCESS_REQUEST_SUCCESS_MESSAGE,
                                    responseType: IAccessRequestResultsType.Success
                                }
                            });
                        }
                    });
                }
            });
        }
    }

    private _requestReportScan(ev?: any) {

        if (this.originalPath) {
            this.setState({ scanRequestSubmissionInProgress: true });
            var reportNumber = this.originalPath.split('/')[this.originalPath.split('/').length - 1];  // Name of Document Set

            const sharePointSearchService = this.props.serviceScope.consume<ISharePointSearchService>(SharePointSearchService.ServiceKey);
            sharePointSearchService.validateAccessRequest("AccessRequest", reportNumber).then((validateAccessRequestResults: IAccessRequestResults | string) => {
                if (!isEmpty(validateAccessRequestResults)) {
                    if (typeof validateAccessRequestResults === "string") {
                        this.setState({
                            scanRequestLogged: true,
                            scanRequestSubmissionInProgress: false,
                            accessRequestResponse: {
                                message: ACCESS_REQUEST_VALIDATE_FAILURE_MESSAGE,
                                responseType: IAccessRequestResultsType.Failure
                            }
                        });
                    } else {

                        if (this.uniqueId) {
                            const element = (
                                <>
                                    <FontIcon aria-label="CompletedSolid" iconName="CompletedSolid" />
                                    <span style={{ color: "green" }}>Scan Request Status : {validateAccessRequestResults.RequestAccessStatus}</span>
                                </>
                            );
                            ReactDOM.render(element, document.getElementById(this.uniqueId));
                        }

                        this.setState({
                            scanRequestLogged: true,
                            scanRequestSubmissionInProgress: false,
                            accessRequestResponse: {
                                message: format(ACCESS_REQUEST_VALIDATE_SUCCESS_MESSAGE, validateAccessRequestResults.RequestAccessStatus),
                                responseType: IAccessRequestResultsType.Success
                            }
                        });
                    }
                } else {
                    sharePointSearchService.submitAccessRequest("AccessRequest", {}).then((accessRequestResults: IAccessRequestResults | string) => {
                        if (isEmpty(accessRequestResults) || (!isEmpty(accessRequestResults) && typeof accessRequestResults === "string")) {
                            this.setState({
                                scanRequestLogged: true,
                                scanRequestSubmissionInProgress: false,
                                accessRequestResponse: {
                                    message: ACCESS_REQUEST_FAILURE_MESSAGE,
                                    responseType: IAccessRequestResultsType.Failure
                                }
                            });
                        } else {

                            if (this.uniqueId) {
                                const element = (
                                    <>
                                        <FontIcon aria-label="CheckMark" iconName="CheckMark" />
                                        <span style={{ color: "green" }}>Scan Request Status : {RequestAccessStatus.New}</span>
                                    </>
                                );
                                ReactDOM.render(element, document.getElementById(this.uniqueId));
                            }

                            this.setState({
                                scanRequestLogged: true,
                                scanRequestSubmissionInProgress: false,
                                accessRequestResponse: {
                                    message: ACCESS_REQUEST_SUCCESS_MESSAGE,
                                    responseType: IAccessRequestResultsType.Success
                                }
                            });
                        }
                    });
                }
            });
        }
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