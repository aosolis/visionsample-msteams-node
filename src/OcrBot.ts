import * as http from "http";
import * as request from "request";
import * as builder from "botbuilder";
import * as consts from "./constants";
import * as utils from "./utils";
import * as vision from "./VisionApi";
import { Strings } from "./locale/locale";
const uuidv4 = require('uuid/v4');

// =========================================================
// Optical Character Recognition Bot
// =========================================================

export class OcrBot extends builder.UniversalBot {

    private visionApi: vision.VisionApi;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);

        this.visionApi = botSettings.visionApi as vision.VisionApi;

        this.dialog(consts.DialogId.Root, this._onMessage.bind(this));
        _connector.onInvoke(this._onInvoke.bind(this));
    }

    // Handle incoming messages
    private async _onMessage(session: builder.Session) {
        const fileUrl = this.getImageFileAttachmentUrl(session.message);
        if (fileUrl) {
            // Image was attached as a file
            this.returnRecognizedTextAsync(session, () => {
                return this.visionApi.runOcrAsync(fileUrl);
            });
            return;
        }
        
        const inlineImageUrl = this.getInlineImageAttachmentUrl(session.message);
        if (inlineImageUrl) {
            // Image was attached as inline content
            this.returnRecognizedTextAsync(session, async () => {
                let buffer = await this.getInlineImageAttachmentAsync(inlineImageUrl, session);
                return await this.visionApi.runOcrAsync(buffer);
            });
            return;
        }

        if (session.message.text) {
            // Try the text as an image URL
            this.returnRecognizedTextAsync(session, () => {
                return this.visionApi.runOcrAsync(session.message.text);
            });
        } else {
            session.send(Strings.help_message);
        }
    }

    // Handle incoming invokes
    private async _onInvoke(event: builder.IEvent, callback: (err: Error, body: any, status?: number) => void): Promise<void> {
        const eventAsAny = event as any;
        if (eventAsAny.name === "fileConsent/invoke") {
            await this.handleFileConsentResponseAsync(event);
            callback(null, "", 200);
        } else {
            callback(new Error("Unknown invoke type: " + eventAsAny), "");
        }
    }

    // Return a caption for the image
    private async returnRecognizedTextAsync(session: builder.Session, ocrOperation: () => Promise<vision.OcrResult>): Promise<void> {
        try {
            const ocrResult = await ocrOperation();
            this.sendOcrResponse(session, ocrResult);
        } catch (e) {
            const errorMessage = (e.result && e.result.reason) || e.message;
            session.send(Strings.analysis_error, errorMessage);
        }
    }

    // Send the OCR response to the user
    private sendOcrResponse(session: builder.Session, result: vision.OcrResult): void {
        let text = this.getRecognizedText(result);

        if (text.length > 0)
        {

            let resultId = uuidv4();
            session.conversationData.ocrResult = {
                resultId: resultId,
                text: text,
            };
    
            const buffer = new Buffer(text, "utf8");
            let fileUploadRequest: builder.IAttachment = {
                contentType: "application/vnd.microsoft.teams.card.file.consent",
                name: session.gettext(Strings.ocr_file_name),
                content: {
                    description: session.gettext(Strings.ocr_file_description),
                    sizeInBytes: buffer.byteLength,
                    acceptContext: {
                        resultId: resultId,
                    }
                },
            };
            let message = new builder.Message(session)
                .text(Strings.ocr_textfound_message)
                .addAttachment(fileUploadRequest);
            session.send(message);
        } else {
            session.send(Strings.ocr_notextfound_message);
        }
    }

    // Handle file consent response
    private async handleFileConsentResponseAsync(event: builder.IEvent): Promise<void> {
        const session = await utils.loadSessionAsync(this, event);

        const value = (event as any).value;        
        switch (value.action) {
            // User declined upload
            case "decline":
                session.conversationData.ocrResult = null;
                session.send(Strings.ocr_file_upload_declined);
                break;

            // User accepted file
            case "accept":
                const ocrResult = session.conversationData.ocrResult;
                const uploadInfo = value.uploadInfo;

                // Check that this is the active OCR result
                if (!ocrResult || (ocrResult.resultId !== value.context.resultId)) {
                    session.send(Strings.ocr_result_expired);
                    return;
                }

                // Upload the content to the file
                const buffer = new Buffer(ocrResult.text, "utf8");
                const options: request.OptionsWithUrl = {
                    url: uploadInfo.uploadUrl,
                    body: buffer,
                    headers: {
                        "Content-Type": "application/octet-stream",
                        "Content-Range": `bytes 0-${buffer.byteLength-1}/${buffer.byteLength}`,
                    },
                };
                request.put(options, (err, res: http.IncomingMessage, body) => {
                    if (err) {
                        console.error(`Error uploading file: ${JSON.stringify(err)}`);
                        session.send(Strings.ocr_upload_error, err.message);
                    } else if ((res.statusCode === 200) || (res.statusCode === 201)) {
                        const fileAttachment = {
                            contentType: "application/vnd.microsoft.teams.card.file.info",
                            contentUrl: uploadInfo.contentUrl,
                            name: uploadInfo.name,
                            content: {
                                uniqueId: uploadInfo.uniqueId,
                                fileType: uploadInfo.fileType,
                            },
                        };
                        const successMessage = new builder.Message(session)
                            .addAttachment(fileAttachment);
                        session.send(successMessage);
                    } else {
                        console.error(`Upload error. statusCode: ${res.statusCode}, body: ${body}`);
                        session.send(Strings.ocr_upload_error, res.statusMessage);
                    }
                });
                break;
        }
    }

    // Return the text recognized in the OCR operation
    private getRecognizedText(result: vision.OcrResult): string {
        const regions = (result.regions || []).map(region => {
            const lines = (region.lines || []).map(line => {
                const words = (line.words || []).map(word => word.text);
                return words.join(" ");
            });
            return lines.join("\r\n");
        });
        return regions.join("\r\n\r\n");
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
