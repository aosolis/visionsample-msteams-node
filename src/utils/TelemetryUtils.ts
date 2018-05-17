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
    const client = appInsights.defaultClient;
    if (client) {
        properties = properties || {};

        const context = client.context;
        const previousUserId = context.tags[context.keys.userId];

        if (botEvent) {
            let address = botEvent.address;
            properties = {
                correlationId: getCorrelationId(address),
                user: address.user.id,
                tenant: messageUtils.getTenantId(botEvent),
                ...properties,
            };
            context.tags[context.keys.userId] = address.user.id;
        }

        client.trackEvent({ name: eventName, properties: properties });
        context.tags[context.keys.userId] = previousUserId;
    }
}

// Log exception to telemetry
export function trackException(error: Error, properties: any = {}, botEvent?: builder.IEvent): void {
    const client = appInsights.defaultClient;
    if (client) {
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
        client.trackException({ exception: error, properties: properties });
    }
}
