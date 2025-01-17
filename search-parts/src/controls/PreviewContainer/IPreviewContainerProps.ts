import { ISharePointSearchService } from "../../services/searchService/ISharePointSearchService";

export interface IPreviewContainerProps {

   /**
    * The element URL to display (can be the iframe source URL)
    */
   elementUrl: string;

   /**
    * The thumbnail image URL
    */
   previewImageUrl: string;

   /**
    * The HTML element to use as target for the callout
    */
   targetElement: HTMLElement;

   /**
    * Indicates if we need to show the preview
    */
   showPreview: boolean;

   /**
    * The preview type
    */
   previewType: PreviewType;

   /**
    * The search result item
    */
   resultItem?: any;

   /**
    * A sharepoint search service instance
    */
   sharePointSearchService: ISharePointSearchService;
}

export enum PreviewType {
   Document
}
