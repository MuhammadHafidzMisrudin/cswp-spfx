import { BaseDialog, IDialogConfiguration } from "@microsoft/sp-dialog";

import styles from "./DialogDemoWebPartsWebPart.module.scss";

export class CustomDialog extends BaseDialog {
  /**
   * name: getConfig
   * return back Dialog Configuration
   */
  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false // will not blocking other dialog.
    };
  }

  /**
   * name: render
   */
  public render(): void {
    this.domElement.innerHTML = `
        <div class="${styles.dialogDemoWebParts}">
          <div class="${styles.container}">
            <div class="${styles.row}">
              <div class="${styles.column}">
                <div>Custom Dialog</div>
                <div><button class="submitButton ${
                  styles.button
                }">submit</button></div>
            </div>
          </div>
        </div>`;

    this.domElement
      .getElementsByClassName("submitButton")[0]
      .addEventListener("click", () => {
        this.close();
      });
  }
}
