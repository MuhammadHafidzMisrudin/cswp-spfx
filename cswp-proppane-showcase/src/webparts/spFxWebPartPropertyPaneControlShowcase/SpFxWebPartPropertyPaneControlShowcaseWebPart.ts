import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxWebPartPropertyPaneControlShowcase.module.scss';
import { ISpFxWebPartPropertyPaneControlShowcaseWebPartProps } from './ISpFxWebPartPropertyPaneControlShowcaseWebPartProps';

import { propertyPaneBuilder } from "../../services/PropPaneBuilder";

export default class SpFxWebPartPropertyPaneControlShowcaseWebPart extends BaseClientSideWebPart<ISpFxWebPartPropertyPaneControlShowcaseWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.spFxWebPartPropertyPaneControlShowcase}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <p class="ms-font-xl ms-fontColor-white">${escape(this.properties.description)}</p>
              <p class="ms-font-l ms-fontColor-white">${escape('Open the property pane to view all the controls')}</p>
              <ol>
                <li>PropertyPaneButton</li>
                <li>PropertyPaneCheckbox</li>
                <li>PropertyPaneChoiceGroup</li>
                <li>PropertyPaneDropDown</li>
                <li>PropertyPaneLink</li>
                <li>PropertyPaneSlider</li>
                <li>PropertyPaneTextField</li>
                <li>PropertyPaneToggle</li>
              </ul>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return propertyPaneBuilder.getPropertyPaneConfiguration();
  }
}
