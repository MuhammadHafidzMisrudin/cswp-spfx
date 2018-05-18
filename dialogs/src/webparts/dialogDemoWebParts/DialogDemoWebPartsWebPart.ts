import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import { Dialog, IAlertOptions, IPromptOptions } from "@microsoft/sp-dialog";

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
              <div><button class="showPrompt ${
                styles.button
              }">show prompt dialog</button></div>
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

    // get a reference to show prompt button.
    this.domElement
      .getElementsByClassName("showPrompt")[0]
      .addEventListener("click", () => {
        this._showPrompt();
      });
  }

  /* Add code here - event handlers  */

  // Alert button
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

  private _confirmOpen(): boolean {
    const decision: boolean = true;
    console.log("confirm open", decision);
    return decision;
  }

  // Prompt button
  private _showPrompt(): void {
    const options: IPromptOptions = {
      confirmOpen: this._confirmOpen
    };

    Dialog.prompt("what is the Voitanos URL?", options).then(
      (result: string | undefined) => {
        console.log(" ", result);
      }
    );
  }

  /* End - event handlers  */

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
