import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField, IPropertyPaneTextFieldProps
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NasaApolloMissionViewerWebPart.module.scss';
import * as strings from 'NasaApolloMissionViewerWebPartStrings';

import {
  IMission
} from "../../models";
import {
  MissionService
} from "../../services";

export interface INasaApolloMissionViewerWebPartProps {
  description: string;
  selectedMission: string;
}

export default class NasaApolloMissionViewerWebPart extends BaseClientSideWebPart<INasaApolloMissionViewerWebPartProps> {

  // retrieve the current selected mission.
  private selectedMission: IMission;

  protected onInit(): Promise<void>{
    return new Promise<void>(
      (
        resolve: () => void,
        reject: (error: any) => void
      ): void => {
        this.selectedMission = this._getSelectedMission();
        resolve();
      });
  }

  private missionDetailElement: HTMLElement;

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.nasaApolloMissionViewer }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Apollo Mission Viewer</span>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <div id="apolloMissionDetails"></div>
            </div>
          </div>
        </div>
      </div>`;

      // get a reference to a div.
      this.missionDetailElement = document.getElementById('apolloMissionDetails');

      // show mission if found, otherwise show empty.
      if (this.selectedMission) {
        this._renderMissionDetails(this.missionDetailElement, this.selectedMission);
      } else {
        this.missionDetailElement.innerHTML = '';
      }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // get the configuration to build the property pane for the web parts.
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('selectedMission', <IPropertyPaneTextFieldProps>{
                  label: 'Apollo Mission ID to Show'
                }) // add new control to a custom property pane.
              ]
            }
          ]
        }
      ]
    };
  }

  protected get disableReactivePropertyChanges(): boolean{
    return true;
  }

  // changes applied: the apply button generated at the bottom of the property pane.
  // this is the result of making the property pane non-reactive.
  protected onAfterPropertyPaneChangesApplied(): void{
    // update selected mission details on web parts.
    this.selectedMission = this._getSelectedMission();

    // update rendering.
    if (this.selectedMission) {
      this._renderMissionDetails(this.missionDetailElement, this.selectedMission);
    } else {
      this.missionDetailElement.innerHTML = '';
    }
  }

  // use Mission Service to retrieve a mission with a corresponding id.
  private _getSelectedMission(): IMission{
    const selectedMissionId: string = (this.properties.selectedMission) ? this.properties.selectedMission : "AS-506";
    return MissionService.getMission(selectedMissionId);
  }

  // display the specified mission details in the provided DOM element.
  private _renderMissionDetails(element: HTMLElement, mission: IMission): void{
    element.innerHTML = `
    <p class="ms-font-m">
      <span class="ms-fontWeight-semibold">Mission: </span>
      ${escape(mission.name)}
    </p>
    <p class="ms-font-m">
      <span class="ms-fontWeight-semibold">Duration: </span>
      ${escape(this._getMissionTimeline(mission))}
    </p>
    <a href="${mission.wiki_href}" target="_blank" class="${styles.button}">
      <span class="${styles.label}">Learn more about ${escape(mission.name)} </span>
    </a>`;
  }

  // return the duration of the mission.
  private _getMissionTimeline(mission: IMission): string{
    let missionDate = mission.end_date !== '' ? `${mission.launch_date.toString()} - ${mission.end_date.toString()}` : `${mission.launch_date.toString()}`;
    return missionDate;
  }

}
