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

import * as http from "http";
import * as request from "request";
import * as builder from "botbuilder";
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

    // Handle incoming messages
    private async _onMessage(session: builder.Session) {
        const fileUrl = utils.getFirstFileAttachmentUrl(session.message);
        if (fileUrl) {
            // Image was attached as a file
            this.returnImageCaptionAsync(session, () => {
                return this.visionApi.describeImageAsync(fileUrl);
            });
            return;
        }
        
        const inlineImageUrl = utils.getFirstInlineImageAttachmentUrl(session.message);
        if (inlineImageUrl) {
            // Image was attached as inline content
            this.returnImageCaptionAsync(session, async () => {
                const buffer = await utils.getInlineAttachmentContentAsync(inlineImageUrl, session);
                return await this.visionApi.describeImageAsync(buffer);
            });
            return;
        }

        if (session.message.text) {
            // Try the text as an image URL
            this.returnImageCaptionAsync(session, () => {
                return this.visionApi.describeImageAsync(session.message.text.trim());
            });
        } else {
            session.send(Strings.help_message);
        }
    }

    // Return a caption for the image
    private async returnImageCaptionAsync(session: builder.Session, describeOperation: () => Promise<vision.DescribeImageResult>): Promise<void> {
        try {
            const describeResult = await describeOperation();
            session.send(Strings.image_caption_response, describeResult.description.captions[0].text);
        } catch (e) {
            session.send(Strings.analysis_error, (e.result && e.result.message) || e.message);
        }
    }
}
