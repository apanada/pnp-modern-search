import { ServiceKey } from "@microsoft/sp-core-library";
import {
  IExtensibilityLibrary, ILayoutDefinition, IComponentDefinition,
  ILayout, LayoutType, ISuggestionProviderDefinition,
  ISuggestionProvider
} from "@pnp/modern-search-extensibility";
import Handlebars from "handlebars";
import { CustomComponentWebComponent } from "../CustomComponent";
import { EbscoLayoutHandlebars } from "../../layouts/ebsco/EbscoLayoutHandlebars";
import { ShellReportsLayoutHandlebars } from "../../layouts/shellReports/ShellReportsLayoutHandlebars";
import { CustomSuggestionProvider } from "../CustomSuggestionProvider";
import { EbscoHelper } from "../../common/utilities/EbscoHelper";
import { isEmpty, unescape } from "@microsoft/sp-lodash-subset";

export class SearchExtensionLibrary implements IExtensibilityLibrary {
  public getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'Ebsco',
        iconName: 'CustomList',
        key: 'EbscoLayoutHandlebars',
        type: LayoutType.Results,
        templateContent: require('../../layouts/ebsco/ebsco-layout.html'),
        serviceKey: ServiceKey.create<ILayout>('Shell:EbscoLayoutHandlebars', EbscoLayoutHandlebars)
      },
      {
        name: 'Shell Reports',
        iconName: 'ReportLibrary',
        key: 'ShellReportsLayoutHandlebars',
        type: LayoutType.Results,
        templateContent: require('../../layouts/shellReports/shell-reports-layout.html'),
        serviceKey: ServiceKey.create<ILayout>('Shell:ShellReportsLayoutHandlebars', ShellReportsLayoutHandlebars)
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'pnp-custom',
        componentClass: CustomComponentWebComponent
      }
    ];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [
      {
        name: 'Custom Suggestions Provider',
        key: 'CustomSuggestionsProvider',
        description: 'A demo custom suggestions provider from the extensibility library',
        serviceKey: ServiceKey.create<ISuggestionProvider>('Shell:CustomSuggestionsProvider', CustomSuggestionProvider)
      }
    ];
  }

  public registerHandlebarsCustomizations(namespace: typeof Handlebars) {

    // Register custom Handlebars helpers
    // Usage {{getEbscoAuthors items}}
    namespace.registerHelper('getEbscoAuthors', (items: Array<{ name?: string; label?: string; group?: string; data?: string; }>) => {

      const authors = [];
      const authorNode = EbscoHelper.getValueFromKeyValuePairs(items, 'Author');

      if (!isEmpty(authorNode)) {

        const authorHtmlStr = unescape(authorNode);

        if (!isEmpty(authorHtmlStr)) {
          try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(authorHtmlStr, 'text/html');

            if (doc && doc.body && doc.body.children && doc.body.children.length > 0) {
              for (let index = 0; index < doc.body.children.length; index++) {
                const itemNode = doc.body.children[index];
                if (itemNode.nodeName === "SEARCHLINK") {
                  authors.push(itemNode.textContent.trim());
                }
              }
            } else {
              authors.push(authorHtmlStr.trim());
            }
          } catch (err) {
            authors.push(authorHtmlStr.trim());
          }
        }
      }

      return new namespace.SafeString(authors.join('; '));
    });

    // Usage {{getEbscoTitle items}}
    namespace.registerHelper('getEbscoTitle', (items: Array<{ name?: string; label?: string; group?: string; data?: string; }>) => {

      const title = EbscoHelper.getValueFromKeyValuePairs(items, 'Title');

      if (!isEmpty(title)) {

        const titleHtmlStr = unescape(EbscoHelper.replaceHighlight(title));

        if (!isEmpty(titleHtmlStr)) {
          return new namespace.SafeString(titleHtmlStr.trim());
        }
      }

      return "";
    });

    // Usage {{getEbscoAbstract items}}
    namespace.registerHelper('getEbscoAbstract', (items: Array<{ name?: string; label?: string; group?: string; data?: string; }>) => {

      const abstract = EbscoHelper.getValueFromKeyValuePairs(items, 'Abstract');

      if (!isEmpty(abstract)) {

        const abstractHtmlStr = unescape(EbscoHelper.replaceHighlight(abstract));

        if (!isEmpty(abstractHtmlStr)) {
          return new namespace.SafeString(abstractHtmlStr.trim());
        }
      }

      return "";
    });

    // Usage {{getEbscoSource items}}
    namespace.registerHelper('getEbscoSource', (items: Array<{ name?: string; label?: string; group?: string; data?: string; }>) => {

      let sourceTitle = "";

      const journalTitle = EbscoHelper.getValueFromKeyValuePairs(items, 'TitleSource', 'Journal Title');
      const source = EbscoHelper.getValueFromKeyValuePairs(items, 'TitleSource', 'Source');
      const relation = EbscoHelper.getValueFromKeyValuePairs(items, 'NoteTitleSource', 'Relation');
      const originalMaterial = EbscoHelper.getValueFromKeyValuePairs(items, 'NoteTitleSource', 'Original Material');

      if (!isEmpty(journalTitle)) {
        sourceTitle = unescape(EbscoHelper.replaceHighlight(journalTitle));
      } else if (!isEmpty(relation)) {
        sourceTitle = unescape(EbscoHelper.replaceHighlight(relation));
      } else if (!isEmpty(originalMaterial)) {
        sourceTitle = unescape(EbscoHelper.replaceHighlight(originalMaterial));
      } else if (!isEmpty(source)) {
        sourceTitle = EbscoHelper.getSearchLink(source);
      }

      return new namespace.SafeString(sourceTitle.trim());
    });

    // Usage {{getEbscoFullText fullText}}
    namespace.registerHelper('getEbscoFullText', (fullText: { customLinks?: Array<{ url?: string; text?: string; mouseOverText?: string; icon?: string; }>; }) => {

      let fullTextObj = {};

      if (!isEmpty(fullText) && fullText.customLinks && fullText.customLinks.length > 0) {
        fullTextObj = {
          fullTextUrl: !isEmpty(fullText.customLinks[0].url) ? unescape(fullText.customLinks[0].url) : null,
          mouseOverText: !isEmpty(fullText.customLinks[0].mouseOverText) ? unescape(fullText.customLinks[0].mouseOverText) : null,
          fullTextImageUrl: !isEmpty(fullText.customLinks[0].icon) ? unescape(fullText.customLinks[0].icon) : null
        };
      }

      return JSON.parse(JSON.stringify(fullTextObj));
    });
  }
}
