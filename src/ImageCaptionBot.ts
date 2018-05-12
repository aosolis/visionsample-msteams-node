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
