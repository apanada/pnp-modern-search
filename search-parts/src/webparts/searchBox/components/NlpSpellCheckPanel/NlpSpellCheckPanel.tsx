import * as React from 'react';
import { INlpSpellCheckPanelProps } from './INlpSpellCheckPanelProps';
import { Link } from 'office-ui-fabric-react';

export default class NlpSpellCheckPanel extends React.Component<INlpSpellCheckPanelProps, null> {

    private _onLinkClick = (queryText: string) => {
        this.setState({
            hidePanel: true
        })
        this.props._onSpellCheckCallback(queryText);
    }

    public render(): React.ReactElement<INlpSpellCheckPanelProps> {
        let renderSpellCheckMessage: JSX.Element = null;

        if (this.props.rawResponse) {
            renderSpellCheckMessage =
                <div>
                    <div>
                        <span>Did you mean </span>
                        <Link onClick={() => this._onLinkClick(this.props.rawResponse.enhancedQuery)} label={`Did you spell it wrong, search for ${this.props.rawResponse.enhancedQuery} instead?`}>
                            <span>{this.props.rawResponse.enhancedQuery}</span>
                        </Link>
                        <span>?</span>
                    </div>

                    <div>
                        <span>Do you want results only for </span>
                        <Link onClick={() => this._onLinkClick(this.props.rawQueryText)}>
                            <span>{this.props.rawQueryText}</span>
                        </Link>
                        <span>?</span>
                    </div>
                </div>;
        }

        return (
            <>
                {renderSpellCheckMessage}
            </>
        );
    }
}