import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField } from '@microsoft/sp-property-pane';
import { PropertyPaneToggle } from "@microsoft/sp-property-pane";
import * as strings from 'SearchExtensionLibraryStrings';

/**
 * Ebsco Layout properties
 */
export interface IEbscoLayoutHandlebarsProperties {

    /**
     * Show or hide the file icon
     */
    showFileIcon: boolean;
}

export class EbscoLayoutHandlebars extends BaseLayout<IEbscoLayoutHandlebarsProperties> {

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {

        // Initializes the property if not defined
        this.properties.showFileIcon = this.properties.showFileIcon !== null && this.properties.showFileIcon !== undefined ? this.properties.showFileIcon : true;

        return [
            PropertyPaneToggle('layoutProperties.showFileIcon', {
                label: strings.Layouts.EbscoList.ShowFileIconLabel
            }),
        ];
    }
}