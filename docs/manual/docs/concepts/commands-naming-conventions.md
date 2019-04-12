# Commands naming conventions

This page briefly describes the command naming conventions used in the CLI

Since the Office 365 CLI is meant to be a productivity tool, the aim is for simpler names e.g. `spo site groupify` to be used so they can be memorized easily. Additionally, aligned are the short options across commands for consistency, e.g. `-u` should relate to the `--webUrl` or `--siteUrl` long name options in all commands.

## Common patterns in the commands naming

Common patterns are used in the commands so a name can look like `teams channel add`. The `teams` part of the command name indicates the service which the user will interact with. The `channel` part of the command name indicates the feature of the service which the user will interact with. The `add` part of the command name indicates the action that can be performed against service or feature of the service.

### Common CRUD pattern like `add`, `get`, `list` and `set` across all commands names

The CLI uses common naming for actions a command will perform against Office 365 service. This is an example table with the naming pattern that is applied to most of the commands.

|Command name|Action|
|---|---|
`teams channel add`| Adds new channel to a Microsoft Teams team
`teams channel get`| Gets a Microsoft Teams team channel information 
`teams channel list`| Lists all channels in a Microsoft Teams team
`teams channel set`| Sets a channel setting in a Microsoft Teams team
`teams channel remove`| Removes a channel from a Microsoft Teams team


## Unified shortcuts so they can be memorized easily for optimal use

Unified shortcuts should allow the CLI users to easily memorize shortcut option and guarantee that this shortcut is consistent across all the commands for easy use.

Short option|Asossiated long option
|---|---
-t|--authType, --targetTypes, --to, --type, --title, --fieldTitle, --targetUrl, --listTitle, --sectionTemplate, --text, --webTemplate
-u|--userName, --usageGuidelinesUrl, --appCatalogUrl, --webUrl, --url, --siteUrl
-p|--password, --path, --properties, --filePath, --policy, --clientSideComponentProperties, --pageSize, --parentFolderUrl, --principals, --promoteAs, --selectProperties, --parentWebUrl
-c|--certificateFile, --packageCreatedBy, --classifications, --channelId, --contentTypeId, --clientSideComponentId, --contentType, --clientState, --column, --classification, --content, --comment
-i|--clientId, --grantId, --appId, --id, --templateId, --groupId, --teamId, --position, --hubSiteId, --includeAssociatedSites, --listId, --requestId, --siteDesignId, --runId, --taskId
-r|--resourceId, --role, --origin, --required, --rights, --resource
-s|--scope, --packageSourceEnvironment, --subject, --siteUrl, --sortOrder, --sourceUrl, --pageSize, --systemUpdate, --section, --siteScripts
-n|--displayName, --name, --packageDisplayName, --userName, --notificationUrl, --pageNumber, --pageName
-e|--environment, --enabled, --commandUIExtension, --expirationDateTime
-d|--packageDescription, --description, --displayName, --defaultClassification, --date
-f|--format, --flow, --fieldId, --filter, --folder, --fileUrl, --folderUrl, --forceRefresh, --fields
-m|--mailNickname, --messageId, --previewImageUrl
-l|--logoPath, --listTitle, --location, --logoUrl, --listId, --filter, --layoutType, --lcid, --locale
-g|--guestUsageGuidelinesUrl, --group
-a|--all, --alias, --previewImageAltText
-j|--joined
-v|--toVersion, --value, --version
-h|--hidden, --hideDefaultThemes
-x|--xml, --translateX
-w|--webUrl, --webTemplate
-parserDisabled|
-q|--query
-y|--translateY
-k|--key
-z|--timeZone
