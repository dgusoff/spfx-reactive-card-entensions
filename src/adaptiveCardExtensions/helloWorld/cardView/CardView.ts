import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton,
  IActionArguments
} from '@microsoft/sp-adaptive-card-extension-base';
import { IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState} from '../HelloWorldAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState> {
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
    if (this.state.currentIndex < this.state.emails.length - 1) {
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
    return buttons as [ICardButton] | [ICardButton, ICardButton];
  } 

  public get data(): IPrimaryTextCardParameters {
    if(this.state.emails.length >0 ) {
      return {
        primaryText: "Title: " + this.state.emails[this.state.currentIndex].Title,
        description: "Desc: " + this.state.emails[this.state.currentIndex].Description
      }
      }
      else{
        return {
          primaryText: "Loading",
          description: "Loading"
        }
      }
  }

  public onAction(action: IActionArguments): void {
    if (action.type === 'Submit') {
      const { id } = action.data;
      switch (id) {
        case 'previous': {
          this.setState({ currentIndex: this.state.currentIndex - 1 , currentEmail:this.state.emails[this.state.currentIndex - 1]});
          break;
        }
        case 'next': {
          this.setState({ currentIndex: this.state.currentIndex + 1 , currentEmail:this.state.emails[this.state.currentIndex + 1]});
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
