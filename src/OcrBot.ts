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

        const fileAttachment = utils.getFirstFileAttachment(session.message);
        if (fileAttachment) {
            // Image was attached as a file
            utils.trackScenarioStart("ocr", { imageSource: "file" }, session.message);
            const resultFilename = fileAttachment.name + ".txt";
            this.returnRecognizedTextAsync(session, () => {
                return this.visionApi.runOcrAsync(fileAttachment.downloadUrl);
            }, resultFilename);
            return;
        }
        
        const inlineImageUrl = utils.getFirstInlineImageAttachmentUrl(session.message);
        if (inlineImageUrl) {
            // Image was attached as inline content
            utils.trackScenarioStart("ocr", { imageSource: "inline" }, session.message);
            this.returnRecognizedTextAsync(session, async () => {
                let buffer = await utils.getInlineAttachmentContentAsync(inlineImageUrl, session);
                return await this.visionApi.runOcrAsync(buffer);
            });
            return;
        }

        if (session.message.text) {
            // Try the text as an image URL
            let urlMatch = session.message.text.match(consts.urlRegExp);
            if (urlMatch) {
                utils.trackScenarioStart("ocr", { imageSource: "url" }, session.message);
                this.returnRecognizedTextAsync(session, () => {
                    return this.visionApi.runOcrAsync(urlMatch[0]);
                });
                return;
            }
        }
        
        // Send help message
        utils.trackScenarioStart("unrecognizedInput", {}, session.message);
        if (session.message.address.conversation.conversationType === "personal") {
            session.send(Strings.ocr_help);
        } else {
            session.send(Strings.ocr_help_paste);
        }
    }

    // Handle incoming invokes
    private async handleInvoke(event: builder.IEvent, callback: (err: Error, body: any, status?: number) => void): Promise<void> {
        // Invokes don't go through middleware, so we have to log them specifically
        LogActivityTelemetry.logIncomingActivity(event);

        const eventAsAny = event as any;
        if (eventAsAny.name === "fileConsent/invoke") {
            // Correlate with the previous event
            const value = (event as any).value;
            utils.setCorrelationId(event.address, value.context.correlationId);

            await this.handleFileConsentResponseAsync(event);
            callback(null, "", 200);
        } else {
            callback(new Error("Unknown invoke type: " + eventAsAny), "");
        }
    }

    // Return text recognized in the image
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
                name: filename || session.gettext(Strings.ocr_file_name),
                content: {
                    description: session.gettext(Strings.ocr_file_description),
                    sizeInBytes: buffer.byteLength,
                    acceptContext: {
                        resultId: resultId,
                        correlationId: utils.getCorrelationId(session.message.address),
                    },
                    declineContext: {
                        resultId: resultId,
                        correlationId: utils.getCorrelationId(session.message.address),
                    }
                },
            };

            session.send(Strings.ocr_textfound_message, langs.where("1", result.language).name || result.language);
            session.send(new builder.Message(session).addAttachment(fileUploadRequest));
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
        let addressOfSourceMessage: builder.IChatConnectorAddress = {
            ...event.address,
            id: event.replyToId,
        };

        const value = (event as any).value;
        switch (value.action) {
            // User declined upload
            case "decline":
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
            case "accept":
                const uploadInfo = value.uploadInfo;

                // Send typing indicator while the file is uploaded
                session.sendTyping();
                session.sendBatch();

                // Check that this is the active OCR result
                if (!lastOcrResult || (lastOcrResult.resultId !== value.context.resultId)) {
                    session.send(Strings.ocr_result_expired);
                    utils.trackScenarioStop("ocr_send", { success: true, status: "expired" }, session.message);
                    return;
                }

                // Upload the content to the file
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

                        // Send message with link to the file
                        const fileAttachment = {
                            contentType: "application/vnd.microsoft.teams.card.file.info",
                            contentUrl: uploadInfo.contentUrl,
                            name: uploadInfo.name,
                            content: {
                                uniqueId: uploadInfo.uniqueId,
                                fileType: uploadInfo.fileType,
                            },
                        };
                        session.send(new builder.Message(session).addAttachment(fileAttachment));
                        utils.trackScenarioStop("ocr_send", { success: true, status: "sent" }, session.message);
                    } else {
                        let uploadError: any = new Error(res.statusMessage);
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
