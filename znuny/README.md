# Read-only DynamicField options operation

Install `Kernel/System/GenericInterface/Operation/DynamicField/Options.pm` below the
Znuny home directory (normally `/opt/otrs/Custom/Kernel/...`), preserving its path.

In the existing GenericInterface web service add an operation named
`DynamicFieldOptions` using the backend
`DynamicField::Options`, and map the HTTP route `GET /DynamicField/Options` to it.
Keep the web service's existing session authentication enabled. The request is:

```text
GET /DynamicField/Options?SessionID=<session>&Names=KostenstelleID,AuftragsID
```

The operation accepts only fields with `ObjectType=Ticket` and
`FieldType=Dropdown`. It calls `DynamicFieldGet` and the official backend
`PossibleValuesGet`; it does not modify data, scrape HTML, or query tables.

Example response:

```json
{
  "Fields": [
    {
      "Name": "KostenstelleID",
      "Label": "Kostenstelle",
      "Options": [
        { "Key": "00000", "Value": "-KEINE KOSTENSTELLE-" },
        { "Key": "390200", "Value": "KOMRO TEAMASSISTENZ/BACK OFFICE" }
      ]
    },
    {
      "Name": "AuftragsID",
      "Label": "Auftrag",
      "Options": [
        { "Key": "00000", "Value": "-KEIN AUFTRAG-" },
        { "Key": "10000006114", "Value": "10000006114 Holzvergasung PGW500" }
      ]
    }
  ]
}
```

After installation rebuild the Znuny configuration/cache according to the local
deployment process and set TaskTool's `DynamicField-Options Route` to the mapped
route. If the operation is unavailable, TaskTool automatically uses its configured
`Key=Display text` fallback lists.
