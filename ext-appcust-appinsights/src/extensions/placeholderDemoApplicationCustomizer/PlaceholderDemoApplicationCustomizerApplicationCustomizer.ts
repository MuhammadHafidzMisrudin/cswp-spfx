import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'PlaceholderDemoApplicationCustomizerApplicationCustomizerStrings';
import { escape } from "@microsoft/sp-lodash-subset";

import styles from './ApolloMissionApplicationCustomizer.module.scss';

const LOG_SOURCE: string = 'PlaceholderDemoApplicationCustomizerApplicationCustomizer';

import { IMission } from "../../models";
import { MissionService } from "../../services";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPlaceholderDemoApplicationCustomizerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PlaceholderDemoApplicationCustomizerApplicationCustomizer
  extends BaseApplicationCustomizer<IPlaceholderDemoApplicationCustomizerApplicationCustomizerProperties> {

    private _topPlaceHolder: PlaceholderContent | undefined;
    private _bottomPlaceHolder: PlaceholderContent | undefined;

    // responsible for loading all information that needs to be rendered in placeholder as well as rendering
    @override
    public onInit(): Promise<void> {
      Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

      this._renderPlaceHolder();
      return Promise.resolve();
    }

    private _onDispose(): void {

    }

    private _renderPlaceHolder(): void {
      if (!this._topPlaceHolder) {
        this._topPlaceHolder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
          onDispose: this._onDispose
        });

        if (!this._topPlaceHolder) {
          console.error("The expected placeholder TOP was not found");
          return;
        }

        if (this._topPlaceHolder.domElement) {
          this._topPlaceHolder.domElement.innerHTML = this._getPlaceholderHtml(MissionService.getMission("AS-506"), "Moon Landing");
        }
      }

      if (!this._bottomPlaceHolder) {
        this._bottomPlaceHolder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, {
          onDispose: this._onDispose
        });

        if (!this._bottomPlaceHolder) {
          console.error("The expected placeholder BOTTOM was not found");
          return;
        }

        if (this._bottomPlaceHolder.domElement) {
          this._bottomPlaceHolder.domElement.innerHTML = this._getPlaceholderHtml(MissionService.getMission("AS-512"), "Last Moon Visit");
        }
      }
    }

    /**
     * Create HTML for insertion into a placeholder on the page.
     *
     * @private
     * @param {IMission}        mission       Apollo mission.
     * @param {string}          prefixMessage String to add before body.
     * @returns {string}                      Html string for insertion into placeholder.
     * @memberof SpaceXMissionNewsApplicationCustomizer
     */
    private _getPlaceholderHtml(mission: IMission, prefixMessage: string): string {
      const missionTime: string = `${this._getLocalizedTimeString(new Date(mission.launch_date))}`;

      const placeholderBody: string = `
            <div class="${styles.app}">
              <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.footer}">
                ${escape(prefixMessage)}:  ${escape(mission.name)} on ${escape(missionTime)}
              </div>
            </div>`;

      return placeholderBody;
    }

    /**
     * Creates localized time string of the provided date/time.
     *
     * @private
     * @param {Date} dateTimestamp  Timestamp to convert to localized time.
     * @returns {string}            Localized time string in human readable format.
     * @memberof SpaceXMissionNewsApplicationCustomizer
     */
    private _getLocalizedTimeString(dateTimestamp: Date): string {
      return `${this._getMonthName(dateTimestamp.getMonth())} ${dateTimestamp.getDate()}, ${dateTimestamp.getFullYear()}`;
    }

    /**
     * Returns a month name based on the provided index.
     *
     * @private
     * @param {number} monthIndex   Month number (0-index).
     * @returns {string}            Month name.
     * @memberof SpaceXMissionNewsApplicationCustomizer
     */
    private _getMonthName(monthIndex: number): string {
      const monthNames: string[] = [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December'
      ];
      return monthNames[monthIndex];
    }
}
