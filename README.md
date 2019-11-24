# TableSharp

A simplified tabular data manager with configurable data structure

## Features

- Edit data in a graphical view
- Save/load data into/from JSON format
- Export data into XLSX format
- Customize data properties, i.e. table columns

## Demo

To be updated

## Customize Data Properties

1. Open a solution `tablesharp.sln`
1. Open `Item.json` under a project `itembuilder`
1. Define a top-level array and add `property`s as many as you want
   - See below for structure of `property`
1. Start `itembuilder` and copy a class definition displayed in the console
1. Paste the class definition to `Item.cs` and build `tablesharp`

### Structure of `property`

```jsonc
{
  "Name": "PropertyName", // Column identifier, a member name of a class `Item`
  "Type": {
    "Value": "C# Built-in type", // Tested: bool, int, string
    "Default": "Default value", // Expressed as C# code
    "Multiline": false // If multiline is required for string
  },
  "Display": {
    "Enabled": true, // Show a column in a spreadsheet if set to `true`
    "Header": "Category", // Column name displayed in a sparedsheet
    "Expression": "Category" // Column content expressed in C# code
  }
}
```
