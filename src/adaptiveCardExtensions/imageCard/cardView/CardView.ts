import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton,
  IActionArguments
} from '@microsoft/sp-adaptive-card-extension-base';
import { IImageCardAdaptiveCardExtensionProps, IImageCardAdaptiveCardExtensionState } from '../ImageCardAdaptiveCardExtension';

export class CardView extends BaseImageCardView<IImageCardAdaptiveCardExtensionProps, IImageCardAdaptiveCardExtensionState> {
  /**
   * Buttons will not be visible if card size is 'Medium' with Image Card View.
   * It will support up to two buttons for 'Large' card size.
   */
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    const buttons: ICardButton[] = [];
    if (this.state.currentIndex > 0) {
      buttons.push({
        title: 'Previous',
        action: {
          type: 'Submit',
          parameters: {
            id: 'previous'
          }
        }
      });
    }
    if (this.state.currentIndex < this.state.alerts.length - 1) {
      buttons.push({
        title: 'Next',
        action: {
          type: 'Submit',
          parameters: {
            id: 'next'
          }
        }
      });
    }
    return buttons as [ICardButton] | [ICardButton, ICardButton];;
  }

  public get data(): IImageCardParameters {
    if(this.state.alerts.length >0 ) {
      return {
        primaryText: "Title: " + this.state.alerts[this.state.currentIndex].Title,
        imageUrl: this.state.alerts[this.state.currentIndex].ImageUrl
      }
      }
      else{
        return {
          primaryText: "Loading",
          imageUrl: ""
        }
      }
  }

  public onAction(action: IActionArguments): void {
    if (action.type === 'Submit') {
      const { id } = action.data;
      switch (id) {
        case 'previous': {
          this.setState({ currentIndex: this.state.currentIndex - 1 , currentAlert:this.state.alerts[this.state.currentIndex - 1]});
          break;
        }
        case 'next': {
          this.setState({ currentIndex: this.state.currentIndex + 1 , currentAlert:this.state.alerts[this.state.currentIndex + 1]});
          break;
        }
        case 'default' : {}
             
      }
    }
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://www.bing.com'
      }
    };
  }
}
