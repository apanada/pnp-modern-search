import { isEmpty, unescape } from "@microsoft/sp-lodash-subset";

export class EbscoHelper {

    public static getValueFromKeyValuePairs = (items: Array<{ name?: string; label?: string; group?: string; data?: string; }>, name: string, label?: string) => {
        if (items && items.length > 0) {

            let filteredItems = [];
            if (!isEmpty(name) && !isEmpty(label)) {
                filteredItems = items.filter(i => i.name === name.toString() && i.label === label.toString());
            } else if (!isEmpty(name)) {
                filteredItems = items.filter(i => i.name === name.toString());
            }

            return filteredItems && filteredItems.length > 0 ? filteredItems[0].data : null;
        }

        return null;
    }

    public static replaceHighlight = (str: string) => {
        return str && str.split('&lt;highlight&gt;').join('<strong>').split('&lt;/highlight&gt;').join('</strong>');
    }

    public static getSearchLink = (node: string) => {
        const nodesText = [];

        if (!isEmpty(node)) {
            const nodeHtmlStr = unescape(EbscoHelper.replaceHighlight(node));

            if (!isEmpty(nodeHtmlStr)) {
                try {
                    const parser = new DOMParser();
                    const doc = parser.parseFromString(nodeHtmlStr, 'text/html');

                    if (doc && doc.body && doc.body.children && doc.body.children.length > 0) {
                        for (let index = 0; index < doc.body.children.length; index++) {
                            const itemNode = doc.body.children[index];
                            if (itemNode.nodeName === "SEARCHLINK") {
                                nodesText.push(itemNode.textContent.trim());
                            }
                        }
                    } else {
                        nodesText.push(nodeHtmlStr.trim());
                    }
                } catch (err) {
                    nodesText.push(nodeHtmlStr.trim());
                }
            }
        }

        return nodesText.join(' ');
    }
}