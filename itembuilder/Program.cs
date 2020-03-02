﻿using System;
using System.Collections.Generic;
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

  class Program
  {
    private static readonly string templateMember = "    public {0} {1} {{ get; set; }}";
    private static readonly string templateDictItem = "      {{ \"{0}\", new Property(\"{1}\") }},";
    private static readonly string templateDicItemMultiline = "      {{ \"{0}\", new Property(\"{1}\", InputType.Multiline) }},";
    private static readonly string templateConsructorArg = "{0} {1}";
    private static readonly string templateConstructor = "      {0} = {1};";
    private static readonly string templateCellHeader = "      cells[row, col++] = \"{0}\";";
    private static readonly string templateCellRow = "      cells[row, col++] = {0};";
    private static readonly string templateClass = @"// THIS IS AUTO-GENERATED CLASS DEFINITION
// DO NOT EDIT THIS FILE
// IF YOU ARE NOT AWARE OF WHAT YOU ARE DOING

using System;
using System.Collections.Generic;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace tablesharp
{{
  class Item
  {{
    // Define properties of each item
{0}

    private static int colMax;

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
      Excel.Range start = cells[row, col];
      cells[row, col].EntireRow.Font.Bold = true;
{5}
      colMax = col - 1;
      Excel.Range end = cells[row, col - 1];
      DrawBorder(cells.Range[start, end]);
      return new Tuple<int, int>(row, col);
    }}

    public Tuple<int, int> FillRow(Excel.Range cells, Tuple<int, int> address)
    {{
      int row = address.Item1;
      int col = address.Item2;
      Excel.Range start = cells[row, col];
      Excel.Range end = cells[row, colMax];
      DrawBorder(cells.Range[start, end]);
      cells.Range[start, end].NumberFormat = ""@"";
{6}
      return new Tuple<int, int>(row, col);
    }}

    private static void DrawBorder(Excel.Range range)
    {{
      Excel.Borders borders = range.Borders;
      borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
      borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
      borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
      borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
      borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
      borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
    }}

    // DO NOT EDIT BELOW
    public static void OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
    {{
      Helper.OnAutoGeneratingColumn(sender, e, itemTypes);
    }}
  }}
}}";

    static void Main(string[] args)
    {
      Assembly assembly = Assembly.GetExecutingAssembly();
      Stream stream = assembly.GetManifestResourceStream("itembuilder.Item.json");
      StreamReader streamReader = new StreamReader(stream);
      string json = streamReader.ReadToEnd();
      List<Item> items = JsonSerializer.Deserialize<List<Item>>(json);
      string members = string.Join("\n", items.ConvertAll(item => string.Format(templateMember, item.Type.Value, item.Name)));
      string dictItems = string.Join("\n", items.ConvertAll(item => item.Type.Multiline ?
        string.Format(templateDicItemMultiline, item.Name, item.Display.Header) :
        string.Format(templateDictItem, item.Name, item.Display.Header)
      ));
      string constructorArgs = string.Join(", ", items.ConvertAll(item => string.Format(templateConsructorArg, item.Type.Value, item.Name.ToLower())));
      string constructors = string.Join("\n", items.ConvertAll(item => string.Format(templateConstructor, item.Name, item.Name.ToLower())));
      string constructorsDefault = string.Join("\n", items.ConvertAll(item => string.Format(templateConstructor, item.Name, item.Type.Default)));
      List<Item> itemsToDisplay = items.FindAll(item => item.Display.Enabled);
      string cellsHeader = string.Join("\n", itemsToDisplay.ConvertAll(item => string.Format(templateCellHeader, item.Display.Header)));
      string cellsRow = string.Join("\n", itemsToDisplay.ConvertAll(item => string.Format(templateCellRow, item.Display.Expression)));
      string classDefinition = string.Format(templateClass, members, dictItems, constructorArgs, constructors, constructorsDefault, cellsHeader, cellsRow);
      Console.WriteLine(classDefinition);
    }
  }
}
