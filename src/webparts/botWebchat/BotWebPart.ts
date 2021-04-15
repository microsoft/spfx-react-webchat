import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BotWebPartStrings';
import BotWebchat from './components/botWebchat';
import { IBotProps } from './contracts/IBotProps';
import { IBotWebPartProps } from './contracts/IBotWebPartProps';

export default class BotWebPart extends BaseClientSideWebPart<IBotWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IBotProps> = React.createElement(
      BotWebchat,
      {
        botButtonText: this.properties.botButtonText,
        chatWindowHeaderTitle: this.properties.chatWindowHeaderTitle,
        context: this.context,
        description: this.properties.description,
        botAuthenticationType: this.properties.botAuthenticationType,
        botDirectLineSecret: this.properties.botDirectLineSecret,
        botDirectLineTokenApiUrl: this.properties.botDirectLineTokenApiUrl,
        botTokenApiResourceId: this.properties.botTokenApiResourceId,
        botTokenApiUrl: this.properties.botTokenApiUrl,        
        botAvatarUrl: this.properties.botAvatarUrl,
        avatarSize: this.properties.avatarSize,
        botAvatarInitials: this.properties.botAvatarInitials,
        backgroundColor: this.properties.backgroundColor,
        bubbleBackground: this.properties.bubbleBackground,
        bubbleBorderRadius: this.properties.bubbleBorderRadius,
        bubbleFromUserBackground: this.properties.bubbleFromUserBackground,
        bubbleFromUserBorderRadius: this.properties.bubbleFromUserBorderRadius,
        bubbleFromUserTextColor: this.properties.bubbleFromUserTextColor,
        suggestedActionBackground: this.properties.suggestedActionBackground,
        suggestedActionTextColor: this.properties.suggestedActionTextColor,
        sendBoxTextWrap:this.properties.sendBoxTextWrap,
        hideScrollToEndButton: this.properties.hideScrollToEndButton
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('chatWindowHeaderTitle', {
                  label: strings.ChatWindowHeaderTitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true
                }),
                PropertyPaneTextField('botButtonText', {
                  label: strings.BotButtonTextFieldLabel
                })          
              ]
            },
            {
              groupName: strings.ConnectionGroupName,
              groupFields: [
                PropertyPaneDropdown('botAuthenticationType', {
                  label: strings.ConnectByFieldLabel,
                  options: [
                    { key: 'Custom-API', text: 'AAD Secured Token API'}, 
                    { key: 'DL-API', text: 'Direct Line Token API'}, 
                    { key: 'DL-Secret', text: 'Direct Line Secret'}],
                  selectedKey: 'DL-Secret'
                }),
                PropertyPaneTextField('botDirectLineSecret', {
                  label: strings.DLSecretFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
