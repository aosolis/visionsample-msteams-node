import * as http from "http";
import * as request from "request";
import * as builder from "botbuilder";
import * as consts from "./constants";
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
        let describeImageOperation: () => Promise<vision.DescribeImageResult>;

        const fileUrl = this.getImageFileAttachmentUrl(session.message);
        if (fileUrl) {
            // Image was attached as a file
            this.returnImageCaptionAsync(session, () => {
                return this.visionApi.describeImageByUrlAsync(fileUrl);
            });
            return;
        }
        
        const inlineImageUrl = this.getInlineImageAttachmentUrl(session.message);
        if (inlineImageUrl) {
            // Image was attached as inline content
            this.returnImageCaptionAsync(session, async () => {
                let buffer = await this.getInlineImageAttachmentAsync(inlineImageUrl, session);
                return await this.visionApi.describeImageBufferAsync(buffer);
            });
            return;
        }

        if (session.message.text) {
            // Try the text as an image URL
            this.returnImageCaptionAsync(session, () => {
                return this.visionApi.describeImageByUrlAsync(session.message.text);
            });
        } else {
            session.send(Strings.help_message);
        }
    }

    // Return a caption for the image
    private async returnImageCaptionAsync(session: builder.Session, describeOperation: () => Promise<vision.DescribeImageResult>): Promise<void> {
        try {
            let describeResult = await describeOperation();
            session.send(Strings.image_caption_response, describeResult.description.captions[0].text);
        } catch (e) {
            session.send(Strings.analysis_error, e.message);
        }
    }

    // Get the content url for an image sent as a file attachment.
    // This url can be downloaded directly without additional authentication, but it is time-limited.
    private getImageFileAttachmentUrl(message: builder.IMessage): string {
        const fileAttachment = message.attachments.find(item => item.contentType === "application/vnd.microsoft.teams.file.download.info");
        if (fileAttachment) {
            return fileAttachment.content.downloadUrl;
        }
        return null;
    }

    // Get the content url for an image sent as an inline attachment.
    // This url requires authentication to download; see getInlineImageAttachmentAsync().
    private getInlineImageAttachmentUrl(message: builder.IMessage): string {
        const imageAttachment = message.attachments.find(item => item.contentType === "image/*");
        if (imageAttachment) {
            return imageAttachment.contentUrl;
        }
        return null;
    }

    // Downloads the image sent as an inline attachment.
    private async getInlineImageAttachmentAsync(contentUrl:string, session: builder.Session): Promise<Buffer> {
        let connector = session.connector as builder.ChatConnector;
        let accessToken = await new Promise<string>((resolve, reject) => {
            connector.getAccessToken((err, accessToken) => {
                if (err) {
                    reject(err);
                } else {
                    resolve(accessToken);
                }
            })
        });
        return await new Promise<Buffer>((resolve, reject) => {
            let options = {
                url: contentUrl,
                headers: {
                    "Authorization": `Bearer ${accessToken}`
                },
                encoding: null,
            };
            request.get(options, (err, res: http.IncomingMessage, body: Buffer) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    reject(new Error(res.statusMessage));
                } else {
                    resolve(body);
                }
            });
        });
    }
}
