# TableSharp

A simplified tabular data manager with configurable data structure

## Features

- Edit data in a graphical view
- Save/load data into/from JSON format
- Export data into XLSX format
- Customize data properties, i.e. table columns

## Demo

To be updated

## How To Use

1. Open the solution `tablesharp.sln`
1. Customize item properties. See below for customizing
1. Build `tablesharp` and use it

### Cutomize Item Properties

1. Open `Definition.json` under the project `itembuilder`
1. Define `Flavors` and `Items`. See below for the entire structure
1. Run `itembuilder`
1. `FlavorData.cs` and `Item.cs` will be generated under the project `tablesharp`

#### Structure of `Definition.json`

```jsonc
{
  "Flavors": [
    "NoFlavor",
    "Flavor1"
  ],
  "Items": [
    {
      "Name": "Property",
      "Type": {
        "Value": "C# Built-in type",
        "Default": "Default value",
        "Multiline": false
      },
      "Display": {
        "Enabled": true,
        "Header": "Category",
        "Expression": "Category"
      }
    }
  ]
}
```

- `Flavors`: Array of strings
  - The first flavor is the default flavor
  - You can add more flavors and customize exported result. See `Expression`
- `Items`: Array of `properties`
  - `Property`
    - `Name`: Property name of `Item` class
    - `Type`: Type of the property
      - `Value`: `string`, `int`, `bool` are supported
      - `Default`: Default value
      - `Multiline`: Set to `true` if multiline string is required
    - `Display`: Export options
      - `Enabled`: Set to `true` if this property shall be displayed as a column on a spreadsheet
      - `Header`: Title to display in the header on a spreadsheet
      - `Expression`: Logic to display. See below for detail

#### Expression

Expression is written in C# language.

Displaying a property value as-is:

```
"Expression": "Property"
```

Displaying a property value conditionally based on other properties:

```
"Expression": "Property1.Length == 0 ? \"\" : Property2"
```

A magic expression `IsFor(flavor)` returns `true` if a user:
- selects the default flavor, for all rows
- selects a specific flavor, for a each row

```
"Expression": "IsFor(flavor) ? Property : \"\""
```

You can combine conditions and `IsFor(flavor)` as many as you want
