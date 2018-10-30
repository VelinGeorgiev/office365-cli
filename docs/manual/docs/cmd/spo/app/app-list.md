# spo app list

Lists apps from the tenant app catalog

## Usage

```sh
spo app list [options]
```

## Options

Option|Description
------|-----------
`-c, --scope [scope]`|Target app catalog. `tenant|sitecollection`. Default `tenant`.
`-s, --siteUrl [siteUrl]`|Absolute URL of the site to install the app in. Required if scope is set to `sitecollection`.
`--help`|Output usage information.
`-o, --output [output]`|Output type. `json|text`. Default `text`.
`--verbose`|Runs command with verbose logging.
`--debug`|Runs command with debug logging.

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To list the apps available in the tenant app catalog, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

When using the text output type (default), the command lists only the values of the `Title`, `ID`, `Deployed` and `AppCatalogVersion` properties of the app. When setting the output type to JSON, all available properties are included in the command output.

When targeting the site collection app catalog, always connect to the SharePoint administration site.

## Examples

List all apps available in the tenant app catalog

```sh
spo app list
```

List all apps available in a site collection app catalog

```sh
spo app list --scope sitecollection --siteUrl https://contoso.sharepoint.com/sites/foo
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)