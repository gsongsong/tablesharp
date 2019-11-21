using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace tablesharp
{
  class Item
  {
    public string Category { get; set; }
    public string FieldName { get; set; }
    public string Description { get; set; }
    public int Size { get; set; }
    public bool IsPublic { get; set; }
    public string Comment { get; set; }

    private static bool DataGridConfigured;

    public Item(string category, string fieldName, string description, int size, bool isPublic, string comment)
    {
      Category = category;
      FieldName = fieldName;
      Description = description;
      Size = size;
      IsPublic = isPublic;
      Comment = comment;
    }

    public Item()
    {
      Category = "";
      FieldName = "";
      Description = "";
      Size = int.MinValue;
      IsPublic = false;
      Comment = "";
    }

    public static void OnAutoGeneratingColumn(DataGridAutoGeneratingColumnEventArgs e)
    {
      switch (e.PropertyName)
      {
        case "Description":
          e.Column = Helper.MultilineTextColumn("Description", "Description");
          break;
        case "IsPublic":
          e.Column = Helper.CheckboxColumn("Public", "IsPublic");
          break;
        default:
          break;
      }
      e.Column.CanUserSort = false;
    }

    public static void ConfigureDataGrid(DataGrid dataGrid)
    {
      if (DataGridConfigured) return;
      // TODO
      DataGridConfigured = true;
    }

    public static Tuple<int, int> FillHeader(Excel.Range cells, Tuple<int, int> addr)
    {
      int row = addr.Item1;
      int col = addr.Item2;
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
  }
}
