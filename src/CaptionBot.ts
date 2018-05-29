// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import * as request from "request";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import * as builderExt from "./botbuilder-extensions";
import * as consts from "./constants";
import * as utils from "./utils";
import * as vision from "./VisionApi";
import { Strings } from "./locale/locale";
import { LogActivityTelemetry } from "./middleware/LogActivityTelemetry";

// =========================================================
// Caption Bot
// =========================================================

export class CaptionBot extends builder.UniversalBot {

    private visionApi: vision.VisionApi;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);

        this.visionApi = botSettings.visionApi as vision.VisionApi;

        this.use(
            new LogActivityTelemetry(),
            new msteams.StripBotAtMentions()
        );

        this.dialog(consts.DialogId.Root, this._onMessage.bind(this));
    }

    // Handle incoming messages
    private async _onMessage(session: builder.Session) {
        session.sendTyping();

        // Caption Bot can take an image file in 3 ways:

        // 1) File attachment -- a file picked from OneDrive or uploaded from the computer
        const fileAttachments = builderExt.FileDownloadInfo.filter(session.message.attachments);
        if (fileAttachments && (fileAttachments.length > 0)) {
            // Image was sent as a file attachment
            // downloadUrl is an unauthenticated URL to the file contents, valid for only a few minutes
            utils.trackScenarioStart("caption", { imageSource: "file" }, session.message);
            this.returnImageCaptionAsync(session, () => {
                return this.visionApi.describeImageAsync(fileAttachments[0].content.downloadUrl);
            });
            return;
        }
        
        // 2) Inline image attachment -- an image pasted into the compose box, or selected from the photo library on mobile
        const inlineImageUrl = utils.getFirstInlineImageAttachmentUrl(session.message);
        if (inlineImageUrl) {
            // Image was sent as inline content
            // contentUrl is a url to the file content; the bot's access token is required 
            utils.trackScenarioStart("caption", { imageSource: "inline" }, session.message);
            this.returnImageCaptionAsync(session, async () => {
                const buffer = await utils.getInlineAttachmentContentAsync(inlineImageUrl, session);
                return await this.visionApi.describeImageAsync(buffer);
            });
            return;
        }

        // 3) URL to an image sent in the text of the message
        if (session.message.text) {
            // Try the text as an image URL
            const urlMatch = session.message.text.match(consts.urlRegExp);
            if (urlMatch) {
                utils.trackScenarioStart("caption", { imageSource: "url" }, session.message);
                this.returnImageCaptionAsync(session, () => {
                    return this.visionApi.describeImageAsync(urlMatch[0]);
                });
                return;
            }
        }
        
        // If none of the above match, send a help message with usage instructions 
        utils.trackScenario("unrecognizedInput", {}, session.message);
        if (session.message.address.conversation.conversationType === "personal") {
            session.send(Strings.image_caption_help);
        } else {
            session.send(Strings.image_caption_help_paste);
        }
    }

    // Return a caption for the image
    private async returnImageCaptionAsync(session: builder.Session, describeOperation: () => Promise<vision.DescribeImageResult>): Promise<void> {
        try {
            const describeResult = await describeOperation();
            if (describeResult.description.captions.length > 0) {
                session.send(Strings.image_caption_response, describeResult.description.captions[0].text);
                utils.trackScenarioStop("caption", { success: true, caption: true }, session.message);
            } else {
                session.send(new builder.Message(session)
                    .text(Strings.image_nocaption_response)
                    .textFormat("xml"));        // Suppress markdown transformation, as it interferes with the shrug response ¯\_(ツ)_/¯
                utils.trackScenarioStop("caption", { success: true, caption: false }, session.message);
            }
        } catch (e) {
            const errorMessage = (e.result && e.result.message) || e.message;
            winston.error(`Failed to analyze image: ${errorMessage}`, e);
            session.send(Strings.analysis_error, errorMessage);
            utils.trackScenarioStop("caption", { success: false, error: errorMessage }, session.message);
        }
    }
}
