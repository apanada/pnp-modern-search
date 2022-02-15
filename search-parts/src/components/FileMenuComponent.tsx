import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { Icon, ITheme, IIconStyles, ImageFit, CommandBarButton, IButtonStyles, IButton, IContextualMenuItem, ContextualMenu, RefObject, IContextualMenuItemProps, IImageProps, Image } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

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

    siteUrl?: string;

    uniqueId?: string;

    parentLink?: string;

    /**
     * The current theme settings
     */
    themeVariant?: IReadonlyTheme;
}

export interface IFileMenuState {

}

export class FileMenu extends React.Component<IFileMenuProps, IFileMenuState> {

    constructor(props: IFileMenuProps) {
        super(props);

        this._openDocumentInBrowser = this._openDocumentInBrowser.bind(this);
        this._openDocumentInApp = this._openDocumentInApp.bind(this);
        this._downloadDocument = this._downloadDocument.bind(this);
        this._openDocumentParentFolder = this._openDocumentParentFolder.bind(this);
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
            { key: 'searchThisSite', iconProps: { iconName: 'Search' }, text: 'Search this site' }
        );

        return <div>
            <CommandBarButton
                text="..."
                styles={styles}
                theme={this.props.themeVariant as ITheme}
                menuProps={{
                    shouldFocusOnMount: true,
                    items: [...menuItems],
                }} />
        </div>;
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
}

export class FileMenuWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        let props = this.resolveAttributes();
        const fileMenu = <FileMenu {...props} />;
        ReactDOM.render(fileMenu, this);
    }
}