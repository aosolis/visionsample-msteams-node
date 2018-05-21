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

import * as appInsights  from "applicationinsights";
import * as winston from "winston";

enum SeverityLevel {
    Verbose = 0,
    Information = 1,
    Warning = 2,
    Error = 3,
    Critical = 4,
}

// Winston transport that logs messages to Application Insights
export class ApplicationInsightsTransport extends winston.Transport {

    constructor(options: any) {
        super(options);
    }

    // ### function log (level, msg, [meta], callback)
    // #### @level {string} Level at which to log the message.
    // #### @msg {string} Message to log
    // #### @meta {Object} **Optional** Additional metadata to attach
    // #### @callback {function} Continuation to respond to when complete.
    // Core logging method exposed to Winston. Metadata is optional.
    public log(level: string, msg: string, meta: any, callback: any): void {
        let self: any = this;

        if (self.silent) {
            return callback(null, true);
        }

        // Track trace events
        let severityLevel = this.getSeverityLevel(level);
        appInsights.defaultClient.trackTrace({ message: msg, severity: severityLevel });

        // Track critical exceptions
        if ((severityLevel === SeverityLevel.Error || severityLevel === SeverityLevel.Critical) &&
            meta && (meta instanceof Error)) {
            appInsights.defaultClient.trackException({ exception: meta });
        }

        self.emit("logged");
        callback(null, true);
    };

    // Convert winston string logging level to Application Insights logging level
    private getSeverityLevel(level: string): SeverityLevel {
        switch (level) {
            case "error":
                return SeverityLevel.Error;
            case "warn":
                return SeverityLevel.Warning;
            case "info":
                return SeverityLevel.Information;
            case "verbose":
            case "debug":
            case "silly":
            default:
                return SeverityLevel.Verbose;
        }
    }
}
