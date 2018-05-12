import * as builder from "botbuilder";
import * as storage from "../storage";

// Creates the storage provider for a bot
export function createBotStorage(config: any): builder.IBotStorage {
    switch (config.storage) {
        case "memory":
            return new builder.MemoryBotStorage();

        case "null":
            return new storage.NullBotStorage();

        case "mongoDb":
            return new storage.MongoDbBotStorage(config.mongoDb.collectionName, config.mongoDb.connectionString);

        default:
            throw new Error(`Unknown storage type ${config.storage}`);
    }
}