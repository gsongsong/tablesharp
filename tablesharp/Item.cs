using System;
using System.Collections.Generic;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace tablesharp
{
  class Item
  {
    // Define properties of each item
    public string Category { get; set; }
    public string FieldName { get; set; }
    public string Description { get; set; }
    public int Size { get; set; }
    public bool IsPublic { get; set; }
    public string Comment { get; set; }

    // Define Header and input type for each property
    // Header will be shown in table editing scene and exported spreadsheet
    // If a property shall support multiline text, use `InputType.Multiline`
    // If a property shall support yes/no or true/false, use `InputType.Checkbox`
    // Otherwise, `InputType` argument is not required
    private static readonly Dictionary<string, Property> itemTypes = new Dictionary<string, Property>
    {
      { "Category", new Property("Category") },
      { "FieldName", new Property("Field name") },
      { "Description", new Property("Description", InputType.Multiline) },
      { "Size", new Property("Size") },
      { "IsPublic", new Property("Public") },
      { "Comment", new Property("Comment", InputType.Multiline) },
    };

    // Define item constructor
    public Item(string category, string fieldName, string description, int size, bool isPublic, string comment)
    {
      Category = category;
      FieldName = fieldName;
      Description = description;
      Size = size;
      IsPublic = isPublic;
      Comment = comment;
    }

    // Define default value for each property
    public Item()
    {
      Category = "";
      FieldName = "";
      Description = "";
      Size = int.MinValue;
      IsPublic = false;
      Comment = "";
    }

    public static Tuple<int, int> FillHeader(Excel.Range cells, Tuple<int, int> addr)
    {
      int row = addr.Item1;
      int col = addr.Item2;
      cells[row, col].EntireRow.Font.Bold = true;
      cells[row, col++] = "Category";
      cells[row, col++] = "Field Name";
      cells[row, col++] = "Description";
      cells[row, col] = "Size";
      return new Tuple<int, int>(row, col);
    }

    public Tuple<int, int> FillRow(Excel.Range cells, Tuple<int, int> address)
    {
      int row = address.Item1;
      int col = address.Item2;
      cells[row, col++] = Category;
      cells[row, col++] = IsPublic ? FieldName : "Reserved";
      cells[row, col++] = IsPublic ? Description : "";
      cells[row, col] = Size;
      return new Tuple<int, int>(row, col);
    }

    // DO NOT EDIT BELOW
    public static void OnAutoGeneratingColumn(DataGridAutoGeneratingColumnEventArgs e)
    {
      Helper.OnAutoGeneratingColumn(e, itemTypes);
    }
  }
}
