let appInsights = require("applicationinsights");
let uuidV4 = require("uuid/v4");
import * as builder from "botbuilder";
import * as messageUtils from "./MessageUtils";

// Add correlation id to address
export function addCorrelationId(address: builder.IAddress): void {
    (address as any).correlationId = uuidV4();
}

// Get correlation id from address
export function getCorrelationId(address: builder.IAddress): string {
    return (address as any).correlationId || "";
}

// Log event to telemetry
export function trackEvent(eventName: string, properties: any = {}, botEvent?: builder.IEvent): void {
    if (appInsights.client) {
        properties = properties || {};
        if (botEvent) {
            let address = botEvent.address;
            properties = {
                correlationId: getCorrelationId(address),
                user: address.user.id,
                tenant: messageUtils.getTenantId(botEvent),
                ...properties,
            };
        }
        appInsights.client.trackEvent(eventName, properties);
    }
}

// Log exception to telemetry
export function trackException(error: Error, properties: any = {}, botEvent?: builder.IEvent): void {
    if (appInsights.client) {
        properties = properties || {};
        if (botEvent) {
            let address = botEvent.address;
            properties = {
                correlationId: getCorrelationId(address),
                user: address.user.id,
                tenant: messageUtils.getTenantId(botEvent),
                ...properties,
            };
        }
        appInsights.client.trackException(error, properties);
    }
}
