import * as http from "http";
import * as request from "request";
import * as querystring from "querystring";

// =========================================================
// Azure Vision API
// =========================================================

// Service endpoint paths
const describePath = "vision/v2.0/describe";

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

export class VisionApi {

    constructor(
        private endpoint: string,
        private accessKey: string,
    )
    {
    }

    // Get a description of the image at the given url
    public async describeImageByUrlAsync(imageUrl: string, language?: string, maxCandidates?:number): Promise<DescribeImageResult> {
        return new Promise<DescribeImageResult>((resolve, reject) => {
            let qsp: any = {
                maxCandidates: maxCandidates || 1,
                language: language || "en",
            };
            let options = {
                url: `https://${this.endpoint}/${describePath}?${querystring.stringify(qsp)}`,
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey
                },
                body: {
                    url: imageUrl,
                },
                json: true,
            };
            request.post(options, (err, res: http.IncomingMessage, body) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    let e = new Error(res.statusMessage) as any;
                    e.statusCode = res.statusCode;
                    try {
                        e.result = JSON.parse(body);
                    } catch (parseError) {
                        console.error("Error parsing body: " + parseError);
                    }
                    reject(e);
                } else {
                    resolve(body as DescribeImageResult);
                }
            });
        });
    }

    // Get a description of the given image at the given url
    public async describeImageBufferAsync(imageBuffer: Buffer, language?: string, maxCandidates?:number): Promise<DescribeImageResult> {
        return new Promise<DescribeImageResult>((resolve, reject) => {
            let qsp: any = {
                maxCandidates: maxCandidates || 1,
                language: language || "en",
            };
            let options = {
                url: `https://${this.endpoint}/${describePath}?${querystring.stringify(qsp)}`,
                headers: {
                    "Ocp-Apim-Subscription-Key": this.accessKey,
                    "Content-Type": "application/octet-stream",
                },
                body: imageBuffer,
            };
            request.post(options, (err, res: http.IncomingMessage, body) => {
                if (err) {
                    reject(err);
                } else if (res.statusCode !== 200) {
                    let e = new Error(res.statusMessage) as any;
                    e.statusCode = res.statusCode;
                    try {
                        e.result = JSON.parse(body);
                    } catch (parseError) {
                        console.error("Error parsing body: " + parseError);
                    }
                    reject(e);
                } else {
                    resolve(JSON.parse(body) as DescribeImageResult);
                }
            });
        });
    }
}
