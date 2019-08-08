# spo app synctoteams

Syncs SharePoint Framework solutions package to Microsoft Teams

## Usage

```sh
spo app synctoteams [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|List item id of the solution package in the tenant app catalog site
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Syncs SPFx solution to Microsoft Teams that has list item id of 1 in the tenant app catalog

```sh
spo app synctoteams --id 1
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins#http-request-9)