/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-debugger */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable prefer-const */
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import { override } from "@microsoft/decorators";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";

import Chat from "../Components/Chat/Chat";
import ChatNoPicture from "../Components/ChatNoPicture/ChatNoPicture";
import * as React from "react";
import * as ReactDOM from "react-dom";

import { graphfi, SPFx } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/photos";

import * as strings from "TeamsChatEmbeddedApplicationCustomizerStrings";

import { app } from "@microsoft/teams-js";

export interface ITeamsChatEmbeddedApplicationCustomizerProperties {}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class TeamsChatEmbeddedApplicationCustomizer extends BaseApplicationCustomizer<ITeamsChatEmbeddedApplicationCustomizerProperties> {
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _isEditMode: boolean = false; // Keep track of the edit mode
  private _observer: MutationObserver | null = null;

  @override
  public async onInit(): Promise<void> {
    // Check initial edit mode state
    this._isEditMode = this._checkIfInEditMode();

    if (!this._isEditMode) {
      await this._renderChat();
    }

    // Listen for page state changes
    this.context.application.navigatedEvent.add(this, this._onNavigatedEvent);

    // Set up MutationObserver to detect DOM changes (for edit mode toggle)
    this._setupEditModeObserver();

    return Promise.resolve();
  }

  private _checkIfInEditMode(): boolean {
    // Check URL for edit mode
    if (window.location.href.toLowerCase().indexOf("mode=edit") !== -1) {
      return true;
    }
  }

  private _setupEditModeObserver(): void {
    // Create mutation observer to watch for edit mode changes
    this._observer = new MutationObserver((mutations) => {
      const currentEditMode = this._checkIfInEditMode();

      // Only update if edit mode state has changed
      if (currentEditMode !== this._isEditMode) {
        this._isEditMode = currentEditMode;

        if (!this._isEditMode) {
          this._renderChat().catch(console.error);
        } else {
          this._clearChat();
        }
      }
    });

    // Start observing the document with the configured parameters
    this._observer.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["class"],
    });
  }

  private async _onNavigatedEvent(): Promise<void> {
    const isInEditMode = this._checkIfInEditMode();

    if (isInEditMode !== this._isEditMode) {
      this._isEditMode = isInEditMode;

      if (!this._isEditMode) {
        await this._renderChat();
      } else {
        this._clearChat();
      }
    }
  }

  private async _renderChat(): Promise<void> {
    try {
      //Detect if the SharePoint page is running inside Microsoft Teams
      //If in Microsoft Teams end the execution
      await app.initialize();
      const context = await app.getContext();
      if (context) {
        return;
      }
    } catch (exp) {
      if (!this._bottomPlaceholder) {
        this._bottomPlaceholder =
          this.context.placeholderProvider.tryCreateContent(
            PlaceholderName.Bottom
          );
      }

      if (!this._bottomPlaceholder) {
        console.error("Could not find bottom placeholder");
        return;
      }

      let profilePictureUrl: string;
      //Get User Profile from Microsoft Graph to ensure the most updated profile picture
      //If permission is not granted by the administrator fallback to the classic SharePoint profile picture
      try {
        const graph = graphfi().using(SPFx(this.context));
        const photoValue = await graph.me.photo.getBlob();
        const url = window.URL || window.webkitURL;
        profilePictureUrl = url.createObjectURL(photoValue);
        //Render Chat component with the user profile picture
        const chat = React.createElement(Chat, {
          label: strings.Label,
          userPhoto: profilePictureUrl,
        });
        ReactDOM.render(chat, this._bottomPlaceholder.domElement);
      } catch (exPhoto) {
        //Render Chat component without the user profile picture
        const chatNoPicture = React.createElement(ChatNoPicture, {
          label: strings.Label,
        });
        ReactDOM.render(chatNoPicture, this._bottomPlaceholder.domElement);
      }
    }
  }

  private _clearChat(): void {
    if (this._bottomPlaceholder && this._bottomPlaceholder.domElement) {
      ReactDOM.unmountComponentAtNode(this._bottomPlaceholder.domElement);
    }
  }

  public onDispose(): void {
    // Clean up the observer when the customizer is disposed
    if (this._observer) {
      this._observer.disconnect();
      this._observer = null;
    }

    // Make sure to clean up the bottom placeholder
    this._clearChat();

    super.onDispose();
  }
}
