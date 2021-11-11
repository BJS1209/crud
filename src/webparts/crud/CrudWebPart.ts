import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import {
  BaseClientSideWebPart,
  WebPartContext,
} from "@microsoft/sp-webpart-base";

import * as strings from "CrudWebPartStrings";
import Crud from "./components/Crud";
import { setup } from "@pnp/common";
import { ICrudProps } from "./components/Crud.types";

export interface ICrudWebPartProps {
  description: string;
  spcontext: WebPartContext;
}

export default class CrudWebPart extends BaseClientSideWebPart<ICrudWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ICrudProps> = React.createElement(Crud, {
      description: this.properties.description,
      spcontext: this.context,
    });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    return super.onInit().then(() => {
      setup({
        spfxContext: this.context, // Provide the SharePoint context
        sp: {
          headers: {
            Accept: "application/json;odata=nometadata",
          },
        },
      });
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
