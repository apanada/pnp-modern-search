import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';

/**
 * Custom Layout properties
 */
export interface ICustomLayoutHandlebarsProperties {
    myTextProperty: string;
}

export class CustomlayoutHandlebars extends BaseLayout<ICustomLayoutHandlebarsProperties> {

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {

        // Initializes the property if not defined
        this.properties.myTextProperty  = this.properties.myTextProperty !== null ? this.properties.myTextProperty : "Default value";
 
        return [
            PropertyPaneTextField('layoutProperties.myTextProperty', {
                label: 'A custom layout property',
                placeholder: 'Fill a value'
            })
        ];
    }
}