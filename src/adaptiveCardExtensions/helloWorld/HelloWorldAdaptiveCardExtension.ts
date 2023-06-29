import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseAdaptiveCardExtension } from "@microsoft/sp-adaptive-card-extension-base";
import { CardView } from "./cardView/CardView";
import { QuickView } from "./quickView/QuickView";
import { HelloWorldPropertyPane } from "./HelloWorldPropertyPane";
// import { MSGraphClientV3 } from '@microsoft/sp-http';

import {
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";

export interface IHelloWorldAdaptiveCardExtensionProps {
  title: string;
}

export interface IHelloWorldAdaptiveCardExtensionState {
  emails: any;
  currentIndex: any;
  currentEmail: any;
}

const CARD_VIEW_REGISTRY_ID: string = "HelloWorld_CARD_VIEW";
export const QUICK_VIEW_REGISTRY_ID: string = "HelloWorld_QUICK_VIEW";

export default class HelloWorldAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloWorldPropertyPane;

  public onInit(): Promise<void> {
    this.state = {
      emails: [],
      currentIndex: 0,
      currentEmail: {},
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(
      QUICK_VIEW_REGISTRY_ID,
      () => new QuickView()
    );
    this.getOutlookData();

    return Promise.resolve();
  }

  private async getOutlookData() {
    let requestUrl =
      "https://m365x52195662.sharepoint.com/_api/web/Lists/GetByTitle('Test List')/Items";
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);

    if (response.ok) {
      const items = await response.json();
      console.log(items.value);
      this.setState({
        currentEmail: items.value[this.state.currentIndex]["Title"],
        emails: items.value
      })
    }

    // this.context.msGraphClientFactory.getClient("3").then((client: MSGraphClientV3): void => {
    //   client.api("/me/mailfolders/Inbox/messages").get((error, messages: any) => {
    //     console.log(messages);
    //     this.setState({currentEmail:messages.value[this.state.currentIndex],emails:messages.value});
    //   });
    // });
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HelloWorld-property-pane'*/
      "./HelloWorldPropertyPane"
    ).then((component) => {
      this._deferredPropertyPane = new component.HelloWorldPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
