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

import * as builder from "botbuilder";
import * as constants from "../constants";
import * as utils from "../utils";

// Log telemetry about Bot Framework activities to app insights
export class LogActivityTelemetry implements builder.IMiddlewareMap {

    // Log incoming activity telemetry
    public static logIncomingActivity(event: builder.IEvent): void {
        utils.ensureCorrelationId(event.address);

        let address = event.address as builder.IChatConnectorAddress;
        let payload: any = {
            type: event.type,
            activityId: address.id,
            bot: address.bot.id,
            conversation: address.conversation.id,
            conversationType: address.conversation.conversationType,
        };

        // Log team and channel id if available
        const teamId = utils.getTeamId(event);
        if (teamId) {
            payload.team = teamId;
            payload.channel = utils.getChannelId(event);
        }

        // Log invoke names for incoming invokes
        if (event.type === constants.invokeType) {
            payload.invokeName = event["name"];
        }

        utils.trackEvent(constants.TelemetryEvent.UserActivity, payload, event);
    }

    // Log incoming activity telemetry
    public readonly receive = (event: builder.IEvent, next: Function): void => {
        LogActivityTelemetry.logIncomingActivity(event);

        next();
    }

    // Log outgoing message telemetry
    public readonly send = (event: builder.IEvent, next: Function): void => {
        if (event.type === constants.messageType) {
            const address = event.address as builder.IChatConnectorAddress;
            const payload: any = {
                type: event.type,
                activityId: address.id,
                bot: address.bot.id,
                conversation: address.conversation.id,
                isGroup: address.conversation.isGroup,
            };

            // Log team and channel id if available
            const teamId = utils.getTeamId(event);
            if (teamId) {
                payload.team = teamId;
                payload.channel = utils.getChannelId(event);
            }

            utils.trackEvent(constants.TelemetryEvent.BotActivity, payload, event);
        }

        next();
    }

}
