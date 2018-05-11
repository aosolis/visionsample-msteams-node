let appInsights = require("applicationinsights");
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
        appInsights.client.trackTrace(msg, severityLevel, meta);

        // Track exceptions
        if ((severityLevel === SeverityLevel.Error || severityLevel === SeverityLevel.Critical) &&
            meta && (meta instanceof Error)) {
            appInsights.client.trackException(meta);
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
