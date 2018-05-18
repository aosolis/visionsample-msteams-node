# visionsample-msteams

This sample demonstrates how a bot can send and receive files in Microsoft Teams.

It consists of two bots:
* **Caption Bot** returns a description of pictures sent to it using the [Azure Computer Vison API](https://docs.microsoft.com/en-us/azure/cognitive-services/Computer-vision/Home#generating-descriptions).
* **OCR Bot** finds text in images that it receives using the [Azure Computer Vision API](https://docs.microsoft.com/en-us/azure/cognitive-services/Computer-vision/Home#optical-character-recognition-ocr). It then sends the user a file containing the text that it recognized.

## Receiving files
A user can send files to a bot in two ways:

#### 1) Directly inserting images into compose box
Users have always been able to send images to a bot by directly inserting the image into the compose box, and then sending the message. On desktop, the user has to copy the image *content* and paste it into the compose box. On mobile, there is a button to insert a picture from the photo library.

Images sent to the bot this way are received as message attachments:
```json
{
    "contentType": "image/*",
    "contentUrl": "https://smba.trafficmanager.net/amer-client-ss.msg/v3/attachments/0-cus-d7-e0ee4ec513aecf2e64124d1b11de2878/views/original"
}
``` 
* `contentType` starts with `image/`.
* `contentUrl` is a resource under the Bot Frameworks `/v3/attachments` API.

To get the binary content of the image, issue a `GET` request for the `contentUrl` of the attachment. This URL requires authentication, so be sure to include an `Authorization` header with the value `Bearer <bot_access_token>`. Obtain the access token using `ChatConnector.getAccessToken()` (Node.js) or [`MicrosoftAppCredentials.GetTokenAsync()`](https://docs.botframework.com/en-us/csharp/builder/sdkreference/db/d61/class_microsoft_1_1_bot_1_1_connector_1_1_microsoft_app_credentials.html#ac12485b537fc010eea4bff2954f8f4fc) (C#).

#### 2) Attaching a file to the message
At Microsoft Build 2018, we announced that bots would soon be able to receive file attachments. The user can click on the "Attach" button, pick a file from their OneDrive for Business library, and then send it to the bot. This works for any kind of file, not just images.

Files sent to the bot also appear as attachments:
```json
{
    "contentType": "application/vnd.microsoft.teams.file.download.info",
    "content": {
        "downloadUrl": "https://<onedrive_download_url>",
        "uniqueId": "70D29F4A-05B7-434A-B7E6-B651FBFEF508",
        "fileType": "tif"
    },
    "contentUrl": "https://<onedrive_path>/phototest.tif",
    "name": "phototest.tif"
}
 ``` 
 * `contentType` is `application/vnd.microsoft.teams.file.download.info`.
 * `contentUrl` is a direct link to the file on OneDrive for Business. Note that this is **not** how your bot will access the file.
 * `name` is the name of the file.
 * `content.fileType` is the file type, as deduced from the file name extension.
 * `content.downloadUrl` is a link to download the file. 
 
 To get the contents of the file, issue a `GET` request to the download URL. The link is unauthenticated, so you do not need to attach an `Authorization` header. Note that it is valid for only a few minutes, so your bot must use it immediately. If you will need the file contents later, download the file and store the contents.

File attachments are only supported in 1:1 chats with the bot. If a user in a channel or group chat tries to send the bot a file attachment, the bot will receive the message, but without the file.

## Sending files
We also announced at Build 2018 that bots will also be able to send files to a user in a 1:1 chat.

To send a file from your bot:

#### 1) Ask the user for permission to upload the file
Request permission by sending a file consent card attachment:
```json
{
    "contentType": "application/vnd.microsoft.teams.card.file.consent",
    "name": "result.txt",
    "content": {
        "description": "Text recognized from image",
        "sizeInBytes": 4348,
        "acceptContext": {
            "resultId": "1a1e318d-8496-471b-9612-720ee4b1b592"
        },
        "declineContext": {
            "resultId": "1a1e318d-8496-471b-9612-720ee4b1b592"
        }
    }
}
```
* `contentType` is `application/vnd.microsoft.teams.card.file.consent`.
* `name` is your proposed file name.
* `content.description` is your proposed file description.
* `content.sizeInBytes` is the size of the file in bytes. 
* `content.acceptContext` and `content.declineContext` are the context values that will be sent to the bot if the user accepts or declines, respectively.

The Teams app renders this card as a permission request, with buttons to `Accept` or `Decline`.

#### 2) User accepts or declines consent
When the user presses either option, the bot receives an `invoke` message:
```
{
    "type": "invoke",
    "name": "fileConsent/invoke",
    ...
    "value": {
        "type": "fileUpload",
        "action": "accept",
        "context": {
            "resultId": "82af7356-d2ef-4430-8a48-1ee189ca01db"
        },
        "uploadInfo": {
            "contentUrl": "https://<sharepoint_url>/result.txt",
            "name": "result.txt",
            "uploadUrl": "https://<sharepoint_upload_url>",
            "uniqueId": "03AA04A0-4D33-4760-94C4-FEE498B68FF2",
            "fileType": "txt"
        }
    }
}
```
* `name` is always `fileConsent/invoke`.
* `value.type` is `fileUpload`.
* `value.action` is `accept` if the user allowed the upload, or `decline` otherwise.
* `value.context` is either the `acceptContext` or the `declineContext` in the card, depending on whether the user accepted or declined.
* `value.uploadInfo.name` is the name of the file.
* `value.uploadInfo.fileType` is the file type, taken from the file name.
* `value.uploadInfo.contentUrl` is a direct link to the file.
* `value.uploadInfo.uploadUrl` is the upload url of the file. This points to a OneDrive [upload session](https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_createuploadsession) for the file. To upload the file contents, follow the instructions [here](https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_createuploadsession#upload-bytes-to-the-upload-session).

When setting the contents of the file, you only need `uploadInfo.uploadUrl`. However, remember the other values in `uploadInfo`, as you will need them to return a file chiclet when you are done uploading the file.

#### 3) Send the user a link to the uploaded file (optional)
We recommend sending a link to the uploaded file. You can send a direct link, using `uploadInfo.contentUrl`, or attach a file chiclet to your message.
```
{
    "contentType": "application/vnd.microsoft.teams.card.file.info",
    "contentUrl": "<uploadInfo.contentUrl>",
    "name": "<uploadInfo.name>",
    "content": {
        "uniqueId": "<uploadInfo.uniqueId>",
        "fileType": "<uploadInfo.fileType>",
    }
}
```
* `contentType` is `application/vnd.microsoft.teams.card.file.info`
* The other fields are taken from the `uploadInfo` value received with the `fileConsent/invoke` message.

#### 4) Delete or update the message with the file consent card
After the use has acted on a file consent card, we recommend deleting or updating the message with the file consent card. This keeps the user from clicking "Accept" or "Decline" again.

The activity id of the message with the consent card is in the `replyToId` of the `fileConsent/invoke` message. You can use the corresponding APIs to delete or update the message. (**Note**: The `replyToId` field is currently empty; a fix has been submitted and will be rolling out to Developer Preview soon.)

## Limitations

Sending and receiving files from bots is still in Developer Preview, and not yet generally available. It's also limited to 1:1 chats

| Method | 1:1 chat | Channels | Group chat |
|---|---|---|---|
| Send image inserted directly into compose box | ✔️ | ✔️ | ✔️ |
| Send file as attachment | ✔️ | ❌ | ❌ |
| Receive file from bot | ✔️ | ❌ | ❌ |

You can use the `conversation.conversationType` property of the incoming message to determine the kind of conversation:
```
{
    ...
    "conversation": {
        "conversationType": "personal",
        "id": "a:1XIx0-SONc5j61Vc-...-s2eVC8jS8wD7rntUk9z8"
    }
    ...
}
```
`conversationType` will be `personal` for a 1:1 with the bot, `channel` for channels, and `groupChat` for group chats.