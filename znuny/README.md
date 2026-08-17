# Read-only DynamicField options operation

Install `Kernel/System/GenericInterface/Operation/DynamicField/Options.pm` below the
Znuny home directory (normally `/opt/otrs/Custom/Kernel/...`), preserving its path.

In the existing GenericInterface web service add an operation named
`DynamicFieldOptions` using the backend
`DynamicField::Options`, and map the HTTP route
`GET /Ticket/DynamicField/:FieldName/Options` to it.
Keep the web service's existing session authentication enabled. The request is:

```text
GET /Ticket/DynamicField/KostenstelleID/Options?SessionID=<session>
```

The operation accepts only fields with `ObjectType=Ticket` and
`FieldType=Dropdown`. It calls `DynamicFieldGet` and the official backend
`PossibleValuesGet`; it does not modify data, scrape HTML, or query tables.

Example response:

```json
{
  "Field": {
      "Name": "KostenstelleID",
      "Label": "Kostenstelle",
      "Options": [
        { "Key": "00000", "Value": "-KEINE KOSTENSTELLE-" },
        { "Key": "390200", "Value": "KOMRO TEAMASSISTENZ/BACK OFFICE" }
      ]
  }
}
```

After installation rebuild the Znuny configuration/cache according to the local
deployment process and set TaskTool's `DynamicField-Options Route` to the mapped
route template. TaskTool calls the route once for `KostenstelleID` and once for
`AuftragsID`. If the operation is unavailable, TaskTool uses its configured
`Key=Display text` fallback lists.
