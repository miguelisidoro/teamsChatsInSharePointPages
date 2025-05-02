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

export default class TeamsChatEmbeddedApplicationCustomizer extends BaseApplicationCustomizer<ITeamsChatEmbeddedApplicationCustomizerProperties> {
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _isEditMode: boolean = false;
  private _observer: MutationObserver | null = null;

  @override
  public async onInit(): Promise<void> {
    this._isEditMode = this._checkIfInEditMode();
    if (!this._isEditMode) await this._renderChat();

    // Set up event listeners
    this.context.application.navigatedEvent.add(
      this,
      this._handleEditModeChange
    );
    this._setupEditModeObserver();

    return Promise.resolve();
  }

  private _checkIfInEditMode(): boolean {
    return (
      window.location.href.toLowerCase().indexOf("mode=edit") !== -1 ||
      document.querySelectorAll(".ms-SPCanvas").length > 0
    );
  }

  private _setupEditModeObserver(): void {
    this._observer = new MutationObserver(async () => {
      const currentEditMode = this._checkIfInEditMode();
      if (currentEditMode !== this._isEditMode) {
        await this._handleEditModeChange();
      }
    });

    this._observer.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["class"],
    });
  }

  private async _handleEditModeChange(): Promise<void> {
    const isInEditMode = this._checkIfInEditMode();
    if (isInEditMode === this._isEditMode) return;

    this._isEditMode = isInEditMode;
    if (!this._isEditMode) {
      await this._renderChat();
    } else {
      this._clearChat();
    }
  }

  private async _renderChat(): Promise<void> {
    try {
      // Skip if running in Teams
      await app.initialize();
      if (await app.getContext()) return;
    } catch (exp) {
      // Only create the placeholder if needed
      if (!this._bottomPlaceholder) {
        this._bottomPlaceholder =
          this.context.placeholderProvider.tryCreateContent(
            PlaceholderName.Bottom
          );
        if (!this._bottomPlaceholder) return;
      }

      try {
        // Try to get user photo and render chat with photo
        const graph = graphfi().using(SPFx(this.context));
        const photoValue = await graph.me.photo.getBlob();
        const profilePictureUrl = URL.createObjectURL(photoValue);
        ReactDOM.render(
          React.createElement(Chat, {
            label: strings.Label,
            userPhoto: profilePictureUrl,
          }),
          this._bottomPlaceholder.domElement
        );
      } catch (exPhoto) {
        // Fallback to chat without photo
        ReactDOM.render(
          React.createElement(ChatNoPicture, { label: strings.Label }),
          this._bottomPlaceholder.domElement
        );
      }
    }
  }

  private _clearChat(): void {
    if (this._bottomPlaceholder?.domElement) {
      ReactDOM.unmountComponentAtNode(this._bottomPlaceholder.domElement);
    }
  }

  public onDispose(): void {
    this._observer?.disconnect();
    this._clearChat();
    super.onDispose();
  }
}
