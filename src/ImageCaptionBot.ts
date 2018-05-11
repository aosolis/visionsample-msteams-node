import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import * as moment from "moment";
import { sprintf } from "sprintf-js";
import * as escapeHtml from "escape-html";
import * as consts from "./constants";
import * as utils from "./utils";
import * as vision from "./VisionApi";
import { Strings } from "./locale/locale";

// =========================================================
// Image Caption Bot
// =========================================================

export class ImageCaptionBot extends builder.UniversalBot {

    private visionApi: vision.VisionApi;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);

        this.visionApi = botSettings.visionApi as vision.VisionApi;

        this.dialog(consts.DialogId.Root, this._onMessage.bind(this));
    }

    private async _onMessage(session: builder.Session) {
        const imageUrl = this.getImageUrl(session.message);

        if (imageUrl) {
            try {
                let describeResult = await this.visionApi.describeImageAsync(imageUrl);
                session.send(Strings.image_caption_response, describeResult.description.captions[0].text);
            } catch (e) {
                session.send(Strings.analysis_error, e.message);
            }
        } else {
            session.send(Strings.help_message);
        }
    }

    private getImageUrl(message: builder.IMessage): string {
        const fileAttachment = message.attachments.find(item => item.contentType === "application/vnd.microsoft.teams.file.download.info");
        if (fileAttachment) {
            return fileAttachment.content.downloadUrl;
        }

        return message.text;
    }
}
