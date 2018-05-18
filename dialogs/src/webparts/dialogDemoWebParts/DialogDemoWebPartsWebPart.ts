import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import { Dialog, IAlertOptions } from "@microsoft/sp-dialog";

import styles from "./DialogDemoWebPartsWebPart.module.scss";
import * as strings from "DialogDemoWebPartsWebPartStrings";

export interface IDialogDemoWebPartsWebPartProps {
  description: string;
}

export default class DialogDemoWebPartsWebPart extends BaseClientSideWebPart<
  IDialogDemoWebPartsWebPartProps
> {
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.dialogDemoWebParts}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">SPFx Dialog Demo</span>
              <div><button class="showAlert ${
                styles.button
              }">show alert dialog</button></div>
            </div>
          </div>
        </div>
      </div>`;

    // get a reference to show alert button.
    this.domElement
      .getElementsByClassName("showAlert")[0]
      .addEventListener("click", () => {
        this._showAlert();
      });
  }

  /* Add code here - event handlers  */

  private _showAlert(): void {
    const options: IAlertOptions = {
      confirmOpen: this._confirmOpen
    };

    Dialog.alert("Congrats, you clicked the alert button.", options).then(
      () => {
        console.log("alert dialog closed");
      }
    );
  }

  /* End - event handlers  */

  private _confirmOpen(): boolean {
    const decision: boolean = true;
    console.log("confirm open", decision);
    return decision;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
