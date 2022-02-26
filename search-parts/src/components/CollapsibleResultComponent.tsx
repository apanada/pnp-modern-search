import * as React from 'react';
import { BaseWebComponent } from '@pnp/modern-search-extensibility';
import * as ReactDOM from 'react-dom';
import { Icon } from 'office-ui-fabric-react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import styles from './CollapsibleContentComponent.module.scss';
import 'core-js/features/dom-collections';
import * as DOMPurify from 'dompurify';

export interface ICollapsibleResultComponentProps {


    /**
     * If the group should be collapsed by default
     */
    defaultCollapsed?: boolean;

    /**
     * Content of the items template
     */
    contentTemplate: string;

    /**
     * The current theme settings
     */
    themeVariant?: IReadonlyTheme;
}

export interface ICollapsibleResultComponentState {

    /**
     * Current collapse/expand state for the group
     */
    isCollapsed: boolean;
}

export class CollapsibleResultComponent extends React.Component<ICollapsibleResultComponentProps, ICollapsibleResultComponentState> {

    private componentRef = React.createRef<HTMLDivElement>();
    private _domPurify: any;

    public constructor(props) {
        super(props);

        this.state = {
            isCollapsed: props.defaultCollapsed ? true : false,
        };

        this._onTogglePanel = this._onTogglePanel.bind(this);

        this._domPurify = DOMPurify.default;
    }


    public render() {

        return <div ref={this.componentRef} data-is-scrollable={true}>
            {
                !this.state.isCollapsed &&
                <div className={styles.collapsible__filterPanel__body__group} style={{ margin: "0px 58px", paddingBottom: "10px" }}>
                    <div dangerouslySetInnerHTML={{ __html: this._domPurify.sanitize(this.props.contentTemplate) }}></div>

                </div>
            }
            <div
                className={styles.collapsible__filterPanel__body__group__collapsibleSection}
                onClick={() => {
                    this._onTogglePanel();
                }}

                onKeyPress={(e) => {
                    if (e.charCode === 13) {
                        this._onTogglePanel();
                    }
                }}
            >
                {
                    this.state.isCollapsed ?
                        <Icon iconName='ChevronDownMed' className={styles.collapsible__filterPanel__body__group__toggleIcon} />
                        :
                        <Icon iconName='ChevronUpMed' className={styles.collapsible__filterPanel__body__group__toggleIcon} />
                }
            </div>
        </div>;
    }

    private _onTogglePanel() {
        this.setState({
            isCollapsed: !this.state.isCollapsed
        });
    }
}

export class CollapsibleResultWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public async connectedCallback() {

        const domParser = new DOMParser();
        const htmlContent: Document = domParser.parseFromString(this.innerHTML, 'text/html');

        // Get the templates
        const contentTemplateContent = htmlContent.getElementById('collapsible-content');

        let contentTemplate = null;

        if (contentTemplateContent) {
            contentTemplate = contentTemplateContent.innerHTML;
        }

        let props = this.resolveAttributes();
        const collapsibleContent = <CollapsibleResultComponent
            {...props}
            contentTemplate={contentTemplate}
        />;

        ReactDOM.render(collapsibleContent, this);
    }
}