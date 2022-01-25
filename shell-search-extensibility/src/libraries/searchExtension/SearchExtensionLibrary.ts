import { ServiceKey } from "@microsoft/sp-core-library";
import {
  IExtensibilityLibrary, ILayoutDefinition, IComponentDefinition,
  ILayout, LayoutType, ISuggestionProviderDefinition,
  ISuggestionProvider
} from "@pnp/modern-search-extensibility";
import { MyCustomComponentWebComponent } from "../CustomComponent";
import { CustomlayoutHandlebars } from "../CustomLayoutHandlebars";
import { CustomSuggestionProvider } from "../CustomSuggestionProvider";

export class SearchExtensionLibrary implements IExtensibilityLibrary {
  public getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'Custom Handlebars',
        iconName: 'Color',
        key: 'CustomLayoutHandlebars',
        type: LayoutType.Results,
        templateContent: require('../custom-layout.html'),
        serviceKey: ServiceKey.create<ILayout>('MyCompany:CustomLayoutHandlebars', CustomlayoutHandlebars)
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'my-custom-component',
        componentClass: MyCustomComponentWebComponent
      }
    ];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [
      {
        name: 'Custom Suggestions Provider',
        key: 'CustomSuggestionsProvider',
        description: 'A demo custom suggestions provider from the extensibility library',
        serviceKey: ServiceKey.create<ISuggestionProvider>('MyCompany:CustomSuggestionsProvider', CustomSuggestionProvider)
      }
    ];
  }

  public registerHandlebarsCustomizations(namespace: typeof Handlebars) {

    // Register custom Handlebars helpers
    // Usage {{myHelper 'value'}}
    namespace.registerHelper('myHelper', (value: string) => {
      return new namespace.SafeString(value.toUpperCase());
    });
  }
}
