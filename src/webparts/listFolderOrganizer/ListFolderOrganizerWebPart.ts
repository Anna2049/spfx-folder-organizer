import * as React from "react";
import * as ReactDom from "react-dom";
import {
  BaseClientSideWebPart,
} from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";

import ListFolderOrganizer from "./components/ListFolderOrganizer";
import { IListFolderOrganizerProps } from "./components/IListFolderOrganizerProps";

export interface IListFolderOrganizerWebPartProps {
  description: string;
}

export default class ListFolderOrganizerWebPart extends BaseClientSideWebPart<IListFolderOrganizerWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IListFolderOrganizerProps> =
      React.createElement(ListFolderOrganizer, {
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
      });

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
            description: "List Folder Organizer Settings",
          },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField("description", {
                  label: "Description",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
