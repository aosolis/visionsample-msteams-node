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
import * as winston from "winston";
import * as langs from "langs";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as consts from "./constants";
import * as utils from "./utils";
import * as vision from "./VisionApi";
import { Strings } from "./locale/locale";
import { LogActivityTelemetry } from "./middleware/LogActivityTelemetry";
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

        this.use(
            new LogActivityTelemetry(),
        );

        this.dialog(consts.DialogId.Root, this.handleMessage.bind(this));
        _connector.onInvoke(this.handleInvoke.bind(this));
    }

    // Handle incoming messages
    private async handleMessage(session: builder.Session) {
        session.sendTyping();

        // OCR Bot can take an image file in 3 ways:

        // 1) File attachment -- a file picked from OneDrive or uploaded from the computer
        const fileAttachments = msteams.FileDownloadInfo.filter(session.message.attachments);
        if (fileAttachments && (fileAttachments.length > 0)) {
            // Image was sent as a file attachment
            // downloadUrl is an unauthenticated URL to the file contents, valid for only a few minutes
            utils.trackScenarioStart("ocr", { imageSource: "file" }, session.message);
            const resultFilename = fileAttachments[0].name + ".txt";
            this.returnRecognizedTextAsync(session, () => {
                return this.visionApi.runOcrAsync(fileAttachments[0].content.downloadUrl);
            }, resultFilename);
            return;
        }
        
        // 2) Inline image attachment -- an image pasted into the compose box, or selected from the photo library on mobile
        const inlineImageUrl = utils.getFirstInlineImageAttachmentUrl(session.message);
        if (inlineImageUrl) {
            // Image was sent as inline content
            // contentUrl is a url to the file content; the bot's access token is required 
            utils.trackScenarioStart("ocr", { imageSource: "inline" }, session.message);
            this.returnRecognizedTextAsync(session, async () => {
                const buffer = await utils.getInlineAttachmentContentAsync(inlineImageUrl, session);
                return await this.visionApi.runOcrAsync(buffer);
            });
            return;
        }

        // 3) URL to an image sent in the text of the message
        if (session.message.text) {
            // Try the text as an image URL
            const urlMatch = session.message.text.match(consts.urlRegExp);
            if (urlMatch) {
                utils.trackScenarioStart("ocr", { imageSource: "url" }, session.message);
                this.returnRecognizedTextAsync(session, () => {
                    return this.visionApi.runOcrAsync(urlMatch[0]);
                });
                return;
            }
        }
        
        // If none of the above match, send a help message with usage instructions 
        utils.trackScenarioStart("unrecognizedInput", {}, session.message);
        if (session.message.address.conversation.conversationType === "personal") {
            session.send(Strings.ocr_help);
        } else {
            session.send(Strings.ocr_help_paste);
        }
    }

    // Handle incoming invoke activities
    private async handleInvoke(event: builder.IEvent, callback: (err: Error, body: any, status?: number) => void): Promise<void> {
        // Invokes don't go through middleware, so we have to log them specifically
        LogActivityTelemetry.logIncomingActivity(event);

        const eventAsAny = event as any;
        if (eventAsAny.name === msteams.fileConsentInvokeName) {
            // Correlate with the previous event
            const value = (event as any).value;
            utils.setCorrelationId(event.address, value.context.correlationId);

            await this.handleFileConsentResponseAsync(event);
            callback(null, "", 200);
        } else {
            callback(new Error("Unknown invoke type: " + eventAsAny), "");
        }
    }

    // Return the result of running OCR on the image
    private async returnRecognizedTextAsync(session: builder.Session, ocrOperation: () => Promise<vision.OcrResult>, filename?: string): Promise<void> {
        try {
            const ocrResult = await ocrOperation();
            this.sendOcrResponse(session, ocrResult, filename);
            utils.trackScenarioStop("ocr", { success: true, ocrLanguage: ocrResult.language }, session.message);
        } catch (e) {
            const errorMessage = (e.result && e.result.message) || e.message;
            winston.error(`Failed to analyze image: ${errorMessage}`, e);
            session.send(Strings.analysis_error, errorMessage);
            utils.trackScenarioStop("ocr", { success: false, error: errorMessage }, session.message);
        }
    }

    // Send the OCR response to the user
    private sendOcrResponse(session: builder.Session, result: vision.OcrResult, filename?: string): void {
        const text = this.getRecognizedText(result);
        if (text.length > 0)
        {
            // Text was found in the message. Send it back to the user as a file.

            // Save the OCR result in conversationData, while we wait for the user's consent to upload the file.
            const resultId = uuidv4();
            session.conversationData.ocrResult = {
                resultId: resultId,
                text: text,
            };
    
            // Calculate the file size in bytes. Note that this only needs to be approximate (upper bound), 
            // so if it's expensive to determine the file size, you don't need to do that.
            // In this case it's straightforward to get the actual size, so we might as well. 
            const buffer = new Buffer(text, "utf8");
            const fileSizeInBytes = buffer.byteLength;

            // Build the file upload consent card
            // Accept and decline context contain: 
            //   1) the result id, to detect the case where the user acted on a stale card
            //   2) a correlation id, so we can correlate the user's response back to this upload consent request.
            const fileConsentCard = new msteams.FileConsentCard(session)
                .name(filename || session.gettext(Strings.ocr_file_name))
                .description(Strings.ocr_file_description)
                .sizeInBytes(fileSizeInBytes)
                .context({
                    resultId: resultId,
                    correlationId: utils.getCorrelationId(session.message.address),
                });

            // Try to convert the language code (e.g., "en", "zh-Hans") from the vision API to a human-readable language name
            const languageCode = result.language.split("-")[0];
            const languageInfo = langs.where("1", languageCode);
            const languageName = (languageInfo && languageInfo.name) || result.language;

            // Send the text prompt and the file consent card.
            // We send them in 2 separate activities, to be sure that we can safely delete the file consent card alone.
            session.send(Strings.ocr_textfound_message, (languageInfo && languageInfo.name) || result.language);
            session.send(new builder.Message(session).addAttachment(fileConsentCard));
            utils.trackScenarioStart("ocr_send", {}, session.message);
        } else {
            session.send(Strings.ocr_notextfound_message);
        }
    }

    // Handle file consent response
    private async handleFileConsentResponseAsync(event: builder.IEvent): Promise<void> {
        const session = await utils.loadSessionAsync(this, event);
        const lastOcrResult = session.conversationData.ocrResult;

        // Create address of source message
        const addressOfSourceMessage: builder.IChatConnectorAddress = {
            ...event.address,
            id: event.replyToId,
        };

        const value = (event as any).value as msteams.IFileConsentCardResponse;
        switch (value.action) {
            // User declined upload
            case msteams.FileConsentCardAction.decline:
                // Delete the file consent card
                if (event.replyToId) {
                    session.connector.delete(addressOfSourceMessage, (err) => {
                        if (err) {
                            winston.error(`Failed to delete consent card: ${err.message}`, err);
                        }
                    });
                }
                session.send(Strings.ocr_file_upload_declined);
                utils.trackScenarioStop("ocr_send", { success: true, status: "declined" }, session.message);
                break;

            // User accepted file
            case msteams.FileConsentCardAction.accept:
                const uploadInfo = value.uploadInfo;

                // Send typing indicator while the file is uploading
                session.sendTyping();
                session.sendBatch();

                // Check that this response is for the the current OCR result
                if (!lastOcrResult || (lastOcrResult.resultId !== value.context.resultId)) {
                    session.send(Strings.ocr_result_expired);
                    utils.trackScenarioStop("ocr_send", { success: true, status: "expired" }, session.message);
                    return;
                }

                // Upload the file contents to the upload session we got from the invoke value
                // See https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_createuploadsession#upload-bytes-to-the-upload-session
                const buffer = new Buffer(lastOcrResult.text, "utf8");
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
                        winston.error(`Error uploading file: ${err.message}`, err);
                        session.send(Strings.ocr_upload_error, err.message);
                        utils.trackScenarioStop("ocr_send", { success: false, error: err.message }, session.message);
                    } else if ((res.statusCode === 200) || (res.statusCode === 201)) {
                        // Delete the file consent card
                        if (event.replyToId) {
                            session.connector.delete(addressOfSourceMessage, (err) => {
                                if (err) {
                                    winston.error(`Failed to delete consent card: ${err.message}`, err);
                                }
                            });
                        }

                        // Send message with link to the file.
                        // The fields in the file info card are populated from the upload info we got in the incoming invoke.
                        const fileInfoCard = msteams.FileInfoCard.fromFileUploadInfo(uploadInfo);
                        session.send(new builder.Message(session).addAttachment(fileInfoCard));
                        utils.trackScenarioStop("ocr_send", { success: true, status: "sent" }, session.message);
                    } else {
                        const uploadError: any = new Error(res.statusMessage);
                        uploadError.body = body;
                        winston.error(`Error uploading file: statusCode:${res.statusCode}`, uploadError);

                        session.send(Strings.ocr_upload_error, uploadError.message);
                        utils.trackScenarioStop("ocr_send", { success: false, error: uploadError.message }, session.message);
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
}
