import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';
import { MSGraphClient } from '@microsoft/sp-http';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import EasyButtonInstructions, { IEasyButtonInstructionsProps } from './components/EasyButtonInstructions';


export default class TeamsCustomizerApplicationCustomizer extends BaseApplicationCustomizer<any> {

  private static headerPlaceholder: PlaceholderContent;
  private cssUrl: string = "https://aptitude4dev.sharepoint.com/sites/PowerPlatformUsers/Shared%20Documents/TeamsConnectedSites.css";

  @override
  public async onInit(): Promise<void> {

    // inject custom css file to hide settings gear icon and home page command bar
    const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
    let customStyle: HTMLLinkElement = document.createElement("link");
    customStyle.href = this.cssUrl;
    customStyle.rel = "stylesheet";
    customStyle.type = "text/css";
    head.insertAdjacentElement("beforeEnd", customStyle);

    const userId = this.context.pageContext.aadInfo.userId._guid;
    const groupId = "d24215a6-496a-4c20-8859-352ed5748c3c"; // "d6ff5742-cb58-47df-9029-8ca2c195f218";
    const graphUrl = `/groups/${groupId}/members`;

    // using MS Graph check if current user is member of IT group
    this.context.application.navigatedEvent.add(this, async () => {
      return await this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api(graphUrl)
            .get(async (error, response: any, rawResponse?: any) => {
              const isGroupMember = await response.value.filter(member => member.id === userId).length > 0 ? true : false;
              console.log(isGroupMember);

              // if current user is member of IT group remove custom css file
              if (this.cssUrl && isGroupMember) {
                const links = document.getElementsByTagName("link");
                console.log(links);
                for (let i = 0; i < links.length; i++) {
                  if (links[i].href.indexOf(this.cssUrl) > -1) {
                    links[i].remove();
                  }
                }
              }

            });
          this.loadReactComponent();
        });

    });
    this._render();
  }

  private _render() {
    if (this.context.placeholderProvider.placeholderNames.indexOf(PlaceholderName.Top) !== -1) {
      if (!TeamsCustomizerApplicationCustomizer.headerPlaceholder || !TeamsCustomizerApplicationCustomizer.headerPlaceholder.domElement) {
        TeamsCustomizerApplicationCustomizer.headerPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
          onDispose: this.onDispose
        });
      }

      this.loadReactComponent();
    }
    else {
      console.log(`The following placeholder names are available`, this.context.placeholderProvider.placeholderNames);
    }
  }

  private loadReactComponent() {
    console.log(this.context.pageContext.list.title);
    if (TeamsCustomizerApplicationCustomizer.headerPlaceholder && TeamsCustomizerApplicationCustomizer.headerPlaceholder.domElement) {
      const element: React.ReactElement<IEasyButtonInstructionsProps> = React.createElement(EasyButtonInstructions, {
        context: this.context
      });

      ReactDom.render(element, TeamsCustomizerApplicationCustomizer.headerPlaceholder.domElement);
    }
    else {
      console.log('DOM element of the header is undefined. Start to re-render.');
      this._render();
    }
  }
}
