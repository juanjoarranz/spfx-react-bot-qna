import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import * as strings from 'BotFrameworkChatWebPartStrings';
import BotFrameworkChat from './components/BotFrameworkChat';
import { IBotFrameworkChatProps } from './components/IBotFrameworkChatProps';

export interface IBotFrameworkChatWebPartProps {
  description: string;
  message: string;
  directLineSecret: string;
  title: string;
  placeholderText: string;
  titleBarBackgroundColor : string;
  botMessagesBackgroundColor: string;
  botMessagesForegroundColor: string;
  userMessagesBackgroundColor: string;
  userMessagesForegroundColor: string;
  messagesRowHeight: string;
}

export default class BotFrameworkChatWebPart extends BaseClientSideWebPart<IBotFrameworkChatWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IBotFrameworkChatProps > = React.createElement(
      BotFrameworkChat,
      {
      description: this.properties.description,
      message: '',
      title: this.properties.title,
      placeholderText: this.properties.placeholderText,
      directLineSecret: this.properties.directLineSecret,
      titleBarBackgroundColor: this.properties.titleBarBackgroundColor,
      botMessagesBackgroundColor: this.properties.botMessagesBackgroundColor,
      botMessagesForegroundColor: this.properties.botMessagesForegroundColor,
      userMessagesBackgroundColor: this.properties.userMessagesBackgroundColor,
      userMessagesForegroundColor: this.properties.userMessagesForegroundColor,
      messagesRowHeight: Number(this.properties.messagesRowHeight),
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Here you can set various properties and settings regarding how your bot chat web part will look visually and functionally work'
          },
          groups: [
            {
              groupName: 'Bot Connection',
              groupFields: [
                PropertyPaneTextField('directLineSecret', {
                  label: 'Direct Line Secret'
                })
              ]
            },
            {
              groupName: 'Appearance',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Title'
                }),
                PropertyPaneTextField('placeholderText', {
                  label: 'Placeholder text'
                } ),
                PropertyPaneSlider('messagesRowHeight', {
                  label: 'Chat Bot Height (pixels)',
                  min: 200,
                  max: 600,
                  step: 10
                } ),
                PropertyFieldColorPicker('titleBarBackgroundColor', {
                  label: 'Title bar background color',
                  selectedColor: this.properties.titleBarBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
                PropertyFieldColorPicker('botMessagesBackgroundColor', {
                  label: 'Bot messages background color',
                  selectedColor: this.properties.botMessagesBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
                PropertyFieldColorPicker('botMessagesForegroundColor', {
                  label: 'Bot messages foreground color',
                  selectedColor: this.properties.botMessagesForegroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                } ),
                PropertyFieldColorPicker('userMessagesBackgroundColor', {
                  label: 'User messages background color',
                  selectedColor: this.properties.userMessagesBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                } ),
                PropertyFieldColorPicker('userMessagesForegroundColor', {
                  label: 'User messages foreground color',
                  selectedColor: this.properties.userMessagesForegroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                } )
              ]
            }
          ]
        }
      ]
    };
  }

  private _validateColorPropertyAsync(value: string): string {
    var colorRegex = /^([a-zA-Z0-9]){6}$/;
    if (!value || colorRegex.test(value) == false) {
      return "Please enter a valid 6 character hex color value";
    }

    return "";
  }
}
