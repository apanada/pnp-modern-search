import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { ITheme, CommandBarButton, IButtonStyles, IContextualMenuItem, IContextualMenuItemProps, IImageProps, Image, Callout, Text, FocusZone, PrimaryButton, DefaultButton, Stack, FocusTrapCallout, mergeStyleSets, FontWeights, FocusZoneTabbableElements } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { UrlHelper } from '../helpers/UrlHelper';
import { ServiceScope } from '@microsoft/sp-core-library';
import { ISharePointSearchService } from '../services/searchService/ISharePointSearchService';
import { SharePointSearchService } from '../services/searchService/SharePointSearchService';

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
     * Image url to use as the icon
     */
    imageUrl?: string;

    redirectUrl?: string;

    originalPath?: string;

    siteId?: string;

    siteUrl?: string;

    uniqueId?: string;

    parentLink?: string;

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
}

const checkAccessStyles = mergeStyleSets({
    callout: {
        width: 320,
        padding: '0px 24px',
    },
    title: {
        marginBottom: 12,
        fontWeight: FontWeights.semilight,
    },
    buttons: {
        display: 'flex',
        justifyContent: 'flex-end',
        marginTop: 20,
    }
});

export class FileMenu extends React.Component<IFileMenuProps, IFileMenuState> {

    constructor(props: IFileMenuProps) {
        super(props);

        this.state = {
            checkAccessCalloutVisible: false,
            userHasAccessToReport: false,
            reportsDocumentSetItemCount: 0
        };

        this._openDocumentInBrowser = this._openDocumentInBrowser.bind(this);
        this._openDocumentInApp = this._openDocumentInApp.bind(this);
        this._downloadDocument = this._downloadDocument.bind(this);
        this._openDocumentParentFolder = this._openDocumentParentFolder.bind(this);
        this._searchThisSite = this._searchThisSite.bind(this);
        this._checkReportAccess = this._checkReportAccess.bind(this);
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

        // Todo: If Content Type === <Report_DocSet_CT>
        menuItems.push({
            key: 'checkReportAccess',
            iconProps: { iconName: 'Signin' },
            text: 'Check access',
            onClick: this._checkReportAccess
        });

        return <div>
            <CommandBarButton
                id='results-menu-item'
                text="..."
                styles={styles}
                theme={this.props.themeVariant as ITheme}
                menuProps={{
                    shouldFocusOnMount: true,
                    items: [...menuItems],
                }} />
            <div>
                {this.state.checkAccessCalloutVisible && (
                    <FocusTrapCallout
                        role="checkaccesscallout"
                        ariaLabelledBy="check-access-callout-label"
                        className={checkAccessStyles.callout}
                        gapSpace={0}
                        onDismiss={() => this.setState({ checkAccessCalloutVisible: false })}
                        target={`#${'results-menu-item'}`}
                        isBeakVisible={false}
                        setInitialFocus
                        styles={{
                            calloutMain: {
                                paddingTop: "20px",
                                paddingBottom: "20px"
                            }
                        }}
                    >
                        <Text block variant="xLarge" className={checkAccessStyles.title}>
                            Reports: Check Permissions
                        </Text>
                        <Text block variant="small">
                            Content is wrapped in a FocusTrapZone so the user cannot accidentally tab or focus out of this callout. Use
                            the buttons to close.
                        </Text>
                        <FocusZone handleTabKey={FocusZoneTabbableElements.all} isCircularNavigation>
                            <Stack className={checkAccessStyles.buttons} gap={8} horizontal>
                                <PrimaryButton onClick={() => this.setState({ checkAccessCalloutVisible: false })}>Done</PrimaryButton>
                                <DefaultButton onClick={() => this.setState({ checkAccessCalloutVisible: false })}>Cancel</DefaultButton>
                            </Stack>
                        </FocusZone>
                    </FocusTrapCallout>
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
            if (this.props.redirectUrl) {
                let documentWebUrl: string = this.props.redirectUrl;
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
            if (this.props.extension && this.props.originalPath) {
                let clientAppUrl: string = `${this._getOfficeClientAppScheme(this.props.extension)}:ofe|u|${this.props.originalPath}`;
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
            if (this.props.siteUrl && this.props.uniqueId) {
                let documentDownloadUrl: string = `${this.props.siteUrl}/_layouts/15/download.aspx?UniqueId=${this.props.uniqueId}`;
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
            if (this.props.parentLink) {
                let parentFolderURl: string = this.props.parentLink;
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
            if (this.props.siteId) {
                let searchThisSiteUrl: string = UrlHelper.addOrReplaceQueryStringParam(window.location.href, "scope", "site");
                searchThisSiteUrl = UrlHelper.addOrReplaceQueryStringParam(searchThisSiteUrl, "sid", this.props.siteId.toString());

                newTabObject = window.open(searchThisSiteUrl, "_blank");
            }
        }
        catch (ex) {
            //optionaly, we can notify the user;
            // cuurently - do nothing
        }
    }

    private _checkReportAccess() {
        this.setState({ checkAccessCalloutVisible: true });

        const sharePointSearchService = this.props.serviceScope.consume<ISharePointSearchService>(SharePointSearchService.ServiceKey);
        const userHasAccess = sharePointSearchService.checkUserAccessToReports("https://m365x083241.sharepoint.com/sites/FlySafeConference/Shell Documents/Shell DocSet");
        userHasAccess.then(hasAccess => {
            this.setState({
                reportsDocumentSetItemCount: hasAccess.ItemCount,
                userHasAccessToReport: hasAccess.hasAccess
            });
        })
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