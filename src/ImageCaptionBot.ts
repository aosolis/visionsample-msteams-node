import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import * as moment from "moment";
import { sprintf } from "sprintf-js";
import * as escapeHtml from "escape-html";
import * as consts from "./constants";
import * as utils from "./utils";
import * as vision from "./VisionApi";
import { Strings } from "./locale/locale";

// =========================================================
// Image Caption Bot
// =========================================================

export class ImageCaptionBot extends builder.UniversalBot {

    private visionApi: vision.VisionApi;

    constructor(
        public _connector: builder.IConnector,
        private botSettings: any,
    )
    {
        super(_connector, botSettings);

        this.visionApi = botSettings.visionApi as vision.VisionApi;

        this.dialog(consts.DialogId.Root, (session) => {
            session.send("Hi!");
        })
    }
}
