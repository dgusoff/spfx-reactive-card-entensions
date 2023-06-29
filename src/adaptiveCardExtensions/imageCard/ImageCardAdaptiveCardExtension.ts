import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { ImageCardPropertyPane } from './ImageCardPropertyPane';

import {
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";

export interface IImageCardAdaptiveCardExtensionProps {
  title: string;
}

export interface IImageCardAdaptiveCardExtensionState {
  alerts: any;
  currentIndex: number;
  currentAlert: any;
}

const CARD_VIEW_REGISTRY_ID: string = 'ImageCard_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'ImageCard_QUICK_VIEW';

export default class ImageCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IImageCardAdaptiveCardExtensionProps,
  IImageCardAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: ImageCardPropertyPane;

  public onInit(): Promise<void> {
    this.state = {
      alerts: [],
      currentIndex: 0,
      currentAlert: {},
     };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    this.getImageData();

    return Promise.resolve();
  }

  private async getImageData(): Promise<void>{
    const requestUrl =
    "https://m365x52195662.sharepoint.com/_api/web/Lists/GetByTitle('Test List')/Items";
  const response: SPHttpClientResponse = await this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

  if (response.ok) {
    const items = await response.json();
    console.log(items.value);
    this.setState({
      currentAlert: items.value[this.state.currentIndex],
      alerts: items.value
    })
  }
  }  

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'ImageCard-property-pane'*/
      './ImageCardPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.ImageCardPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
