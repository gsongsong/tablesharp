﻿using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.Json;

namespace itembuilder
{
  class Display
  {
    public bool Enabled { get; set; }
    public string Expression { get; set; }
    public string Header { get; set; }
  }
  class Type
  {
    public string Default { get; set; }
    public bool Multiline { get; set; }
    public string Value { get; set; }
  }
  class Item
  {
    public string Name { get; set; }
    public Type Type { get; set; }
    public Display Display { get; set; }
  }
  class Definition
  {
    public List<string> Flavors { get; set; }
    public List<Item> Items { get; set; }
  }

  class Program
  {
    private static readonly string templateFlavorItem = @"        ""{0}"",";
    private static readonly string templateMember = "    public {0} {1} {{ get; set; }}";
    private static readonly string templateDictItem = "      {{ \"{0}\", new Property(\"{1}\") }},";
    private static readonly string templateDicItemMultiline = "      {{ \"{0}\", new Property(\"{1}\", InputType.Multiline) }},";
    private static readonly string templateConsructorArg = "{0} {1}";
    private static readonly string templateConstructor = "      {0} = {1};";
    private static readonly string templateCellHeader = "      cells[row, col++] = \"{0}\";";
    private static readonly string templateCellRow = "      cells[row, col++] = {0};";
    private static readonly string templateFlavorDataClass = @"/**
 * FlavorData.cs
 */
// THIS IS AUTO-GENERATED CLASS DEFINITION
// DO NOT EDIT THIS FILE
// IF YOU ARE NOT AWARE OF WHAT YOU ARE DOING

using System.Collections.ObjectModel;

namespace tablesharp
{{
  class FlavorData
  {{
    public ObservableCollection<string> List {{ get; set; }}
    public string Selected {{ get; set; }}

    public FlavorData()
    {{
      List = new ObservableCollection<string>
      {{
{0}
      }};
      Selected = ""{1}"";
    }}
  }}
}}
";
    private static readonly string templateItemClass = @"/**
 * Item.cs
 */
// THIS IS AUTO-GENERATED CLASS DEFINITION
// DO NOT EDIT THIS FILE
// IF YOU ARE NOT AWARE OF WHAT YOU ARE DOING

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace tablesharp
{{
  class Item
  {{
    // Define properties of each item
{0}

    // Define Header and input type for each property
    // Header will be shown in table editing scene and exported spreadsheet
    // If a property shall support multiline text, use `InputType.Multiline`
    // If a property shall support yes/no or true/false, use `InputType.Checkbox`
    // Otherwise, `InputType` argument is not required
    private static readonly Dictionary<string, Property> itemTypes = new Dictionary<string, Property>
    {{
{1}
    }};

    // Define item constructor
    public Item({2})
    {{
{3}
    }}

    // Define default value for each property
    public Item()
    {{
{4}
    }}

    public static Tuple<int, int> FillHeader(Excel.Range cells, Tuple<int, int> addr)
    {{
      int row = addr.Item1;
      int col = addr.Item2;
{5}
      return new Tuple<int, int>(row, col);
    }}

    public Tuple<int, int> FillRow(Excel.Range cells, Tuple<int, int> address, string flavor)
    {{
      int row = address.Item1;
      int col = address.Item2;
{6}
      return new Tuple<int, int>(row, col);
    }}

    // DO NOT EDIT BELOW
    public static void OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
    {{
      Helper.OnAutoGeneratingColumn(sender, e, itemTypes);
    }}

    public bool IsFor(string flavor)
    {{
      PropertyInfo propertyInfo = GetType().GetProperty(flavor);
      if (propertyInfo == null)
      {{
        return true;
      }}
      else
      {{
        return (bool) propertyInfo.GetValue(this, null);
      }}
    }}
  }}
}}
";

    static void Main(string[] args)
    {
      string tablesharpPath = args[0];
      string definitionPath = Path.Combine(tablesharpPath, "Definition.json");
      string flavorDataPath = Path.Combine(tablesharpPath, "FlavorData.cs");
      string itemPath = Path.Combine(tablesharpPath, "Item.cs");

      Assembly assembly = Assembly.GetExecutingAssembly();
      FileStream stream = new FileStream(definitionPath, FileMode.Open);
      StreamReader streamReader = new StreamReader(stream);
      string json = streamReader.ReadToEnd();
      Definition definition = JsonSerializer.Deserialize<Definition>(json);
      List<string> flavors = definition.Flavors;
      List<string> flavorsExceptFirst = flavors.GetRange(1, flavors.Count - 1);
      List<Item> items = definition.Items;

      string flavorsItem = string.Join("\n", definition.Flavors.ConvertAll(flavor => string.Format(templateFlavorItem, flavor)));
      string flavorDataClass = string.Format(templateFlavorDataClass, flavorsItem, definition.Flavors[0]);
      File.WriteAllText(flavorDataPath, flavorDataClass);

      List<string> memberList = new List<string>
      {
        string.Join("\n", items.ConvertAll(item => string.Format(templateMember, item.Type.Value, item.Name))),
        string.Join("\n", flavorsExceptFirst.ConvertAll(flavor => string.Format(templateMember, "bool", flavor))),
      };
      string members = string.Join("\n", memberList);

      List<string> dictItemList = new List<string>
      {
        string.Join("\n", items.ConvertAll(item => item.Type.Multiline ?
          string.Format(templateDicItemMultiline, item.Name, item.Display.Header) :
          string.Format(templateDictItem, item.Name, item.Display.Header)
        )),
        string.Join("\n", flavorsExceptFirst.ConvertAll(flavor => string.Format(templateDictItem, flavor, flavor))),
      };
      string dictItems = string.Join("\n", dictItemList);

      List<string> consturctorArgList = new List<string>
      {
        string.Join(", ", items.ConvertAll(item => string.Format(templateConsructorArg, item.Type.Value, item.Name.ToLower()))),
        string.Join(", ", flavorsExceptFirst.ConvertAll(flavor => string.Format(templateConsructorArg, "bool", flavor.ToLower()))),
      };
      string constructorArgs = string.Join(", ", consturctorArgList);

      List<string> constructorList = new List<string>
      {
        string.Join("\n", items.ConvertAll(item => string.Format(templateConstructor, item.Name, item.Name.ToLower()))),
        string.Join("\n", flavorsExceptFirst.ConvertAll(flavor => string.Format(templateConstructor, flavor, flavor.ToLower()))),
      };
      string constructors = string.Join("\n", constructorList);

      List<string> constructorDefaulList = new List<string>
      {
        string.Join("\n", items.ConvertAll(item => string.Format(templateConstructor, item.Name, item.Type.Default))),
        string.Join("\n", flavorsExceptFirst.ConvertAll(flavor => string.Format(templateConstructor, flavor, "true"))),
      };
      string constructorsDefault = string.Join("\n", constructorDefaulList);

      List<Item> itemsToDisplay = items.FindAll(item => item.Display.Enabled);
      string cellsHeader = string.Join("\n", itemsToDisplay.ConvertAll(item => string.Format(templateCellHeader, item.Display.Header)));
      string cellsRow = string.Join("\n", itemsToDisplay.ConvertAll(item => string.Format(templateCellRow, item.Display.Expression)));
      string itemClass = string.Format(templateItemClass, members, dictItems, constructorArgs, constructors, constructorsDefault, cellsHeader, cellsRow);
      File.WriteAllText(itemPath, itemClass);
    }
  }
}
