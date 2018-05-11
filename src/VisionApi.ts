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

    public async describeImageAsync(imageUrl: string, language?: string, maxCandidates?:number): Promise<DescribeImageResult> {
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
                    reject(new Error(res.statusMessage));
                } else {
                    resolve(body as DescribeImageResult);
                }
            });
        });
    }
}
