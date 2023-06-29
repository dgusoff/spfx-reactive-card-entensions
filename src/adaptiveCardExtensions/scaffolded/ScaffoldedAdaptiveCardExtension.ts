import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { ScaffoldedPropertyPane } from './ScaffoldedPropertyPane';

export interface IScaffoldedAdaptiveCardExtensionProps {
  title: string;
}

export interface IScaffoldedAdaptiveCardExtensionState {
}

const CARD_VIEW_REGISTRY_ID: string = 'Scaffolded_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'Scaffolded_QUICK_VIEW';

export default class ScaffoldedAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IScaffoldedAdaptiveCardExtensionProps,
  IScaffoldedAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: ScaffoldedPropertyPane;

  public onInit(): Promise<void> {
    this.state = { };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'Scaffolded-property-pane'*/
      './ScaffoldedPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.ScaffoldedPropertyPane();
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
