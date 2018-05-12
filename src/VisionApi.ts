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
import * as querystring from "querystring";
import * as winston from "winston";

// =========================================================
// Azure Vision API
// =========================================================

// Service endpoint paths
const describePath = "vision/v2.0/describe";
const ocrPath = "vision/v2.0/ocr";

export interface DescribeImageResult {
    description: ImageDescription;
    requestId: string;
    metadata: ImageMetadata;
}

export interface ImageDescription {
    tags: string[];
    captions: ImageCaption[];
}

export interface ImageCaption {
    text: string;
    confidence: number;
}

export interface ImageMetadata {
    width: number;
    height: number;
    format: string;
}

export interface OcrResult {
    language: string;
    textAngle: number;
    orientation: "Up"|"Down"|"Left"|"Right";
    regions: OcrTextRegion[];
}

export interface OcrTextRegion {
    boundingBox: string;
    lines: OcrTextLine[];
}

export interface OcrTextLine {
    boundingBox: string;
    words: OcrWord[];
}

export interface OcrWord {
    boundingBox: string;
    text: string;
}


export class VisionApi {

    constructor(
        private endpoint: string,
        private accessKey: string,
    )
    {
    }

    // Get a description of the image
    public async describeImageAsync(image: string|Buffer, language?: string, maxCandidates?:number): Promise<DescribeImageResult> {
        return new Promise<DescribeImageResult>((resolve, reject) => {
            let qsp: any = {
                maxCandidates: maxCandidates || 1,
                language: language || "en",
            };
            let options: request.OptionsWithUrl = {
                url: `https://${this.endpoint}/${describePath}?${querystring.stringify(qsp)}`,
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey
                },
            };

            if (typeof(image) === "string") {
                options.body = { url: image };
                options.json = true;
            } else {
                options.headers["Content-Type"] = "application/octet-stream";
                options.body = image;
            }

            request.post(options, (err, res: http.IncomingMessage, body) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    let e = new Error(res.statusMessage) as any;
                    e.statusCode = res.statusCode;
                    try {
                        e.result = this.parseJSONIfString(body);
                    } catch (parseError) {
                        winston.error(`Error parsing body: ${parseError.message}`, body);
                    }
                    reject(e);
                } else {
                    resolve(this.parseJSONIfString(body) as DescribeImageResult);
                }
            });
        });
    }

    // Get a description of the image
    public async runOcrAsync(image: string|Buffer, language?: string, maxCandidates?:number): Promise<OcrResult> {
        return new Promise<OcrResult>((resolve, reject) => {
            let qsp: any = {
                detectOrientation: true,
                language: language || "en",
            };
            let options: request.OptionsWithUrl = {
                url: `https://${this.endpoint}/${ocrPath}?${querystring.stringify(qsp)}`,
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey
                },
            };

            if (typeof(image) === "string") {
                options.body = { url: image };
                options.json = true;
            } else {
                options.headers["Content-Type"] = "application/octet-stream";
                options.body = image;
            }

            request.post(options, (err, res: http.IncomingMessage, body) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    let e = new Error(res.statusMessage) as any;
                    e.statusCode = res.statusCode;
                    try {
                        e.result = this.parseJSONIfString(body);
                    } catch (parseError) {
                        winston.error(`Error parsing body: ${parseError.message}`, body);
                    }
                    reject(e);
                } else {
                    resolve(this.parseJSONIfString(body) as OcrResult);
                }
            });
        });
    }

    // Parse the body as JSON if it is a string
    private parseJSONIfString(body: string|object): any {
        if (typeof(body) === "string") {
            return JSON.parse(body);
        }
        return body;
    }
}
