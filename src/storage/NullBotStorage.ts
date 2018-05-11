// Copyright (c) Microsoft. All rights reserved.

import * as builder from "botbuilder";

/** Replacable storage system used by UniversalBot. */
export class NullBotStorage implements builder.IBotStorage {

    // Reads in data from storage
    public getData(context: builder.IBotStorageContext, callback: (err: Error, data: builder.IBotStorageData) => void): void {
        callback(null, {});
    }

    // Writes out data from storage
    public saveData(context: builder.IBotStorageContext, data: builder.IBotStorageData, callback?: (err: Error) => void): void {
        callback(null);
    }
}
