// Telemetry processor that strips query parameters from urls
// Query parameters can potentially contain PII, so it's safer to strip them out
export function stripQueryFromTelemetryUrls(envelope: any, context: any): void {
    if ((envelope.data.baseType === "RemoteDependencyData") && (envelope.data.baseData.type === "Http")) {
        let url = envelope.data.baseData.data;
        envelope.data.baseData.data = stripQueryFromUrl(url);
    }
}

// Strip query part from a url
function stripQueryFromUrl(url: string): string {
    let firstQuestionMarkIndex = url.indexOf("?");
    if (firstQuestionMarkIndex >= 0) {
        return url.substring(0, firstQuestionMarkIndex);
    } else {
        return url;
    }
}
