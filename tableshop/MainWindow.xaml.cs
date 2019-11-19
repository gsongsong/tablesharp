using Microsoft.Win32;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace tableshop
{
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : Window
  {
    private ObservableCollection<Item> dataTable;
    public MainWindow()
    {
      InitializeComponent();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      dataTable = new ObservableCollection<Item>();
      dataGrid.DataContext = dataTable;
    }

    private void Open(object sender, RoutedEventArgs e)
    {
      OpenFileDialog openFileDialog = new OpenFileDialog
      {
        DefaultExt = ".json",
        Filter = "JSON files (.json)|*.json",
      };
      bool? dialogResult = openFileDialog.ShowDialog();
      if (dialogResult == true)
      {
        filename.Content = openFileDialog.FileName;
        string json = File.ReadAllText(openFileDialog.FileName);
        List<Item> list = JsonSerializer.Deserialize<List<Item>>(json);
        dataTable = new ObservableCollection<Item>(list);
        dataGrid.DataContext = dataTable;
      }
    }

    private void SaveAs(object sender, RoutedEventArgs e)
    {
      SaveFileDialog saveFileDialog = new SaveFileDialog
      {
        DefaultExt = ".json",
        Filter = "JSON files (.json)|*.json",
      };
      bool? dialogResult = saveFileDialog.ShowDialog();
      if (dialogResult == true)
      {
        string json = JsonSerializer.Serialize(dataTable, new JsonSerializerOptions
        {
          WriteIndented = true,
        });
        File.WriteAllText(saveFileDialog.FileName, json);
      }
    }

    private void Export(object sender, RoutedEventArgs e)
    {
      SaveFileDialog saveFileDialog = new SaveFileDialog
      {
        DefaultExt = ".xlsx",
        Filter = "Excel files (.xlsx)|*.xlsx",
      };
      bool? dialogResult = saveFileDialog.ShowDialog();
      if (dialogResult == true)
      {
        var app = new Excel.Application();
        app.Workbooks.Add();
        Excel._Worksheet ws = app.ActiveSheet;
        Excel.Range cells = ws.Cells;
        int row = 1;
        int col = 1;
        cells[row, col++] = "Category";
        cells[row, col++] = "Field Name";
        cells[row, col++] = "Description";
        cells[row, col++] = "Size";
        foreach (Item item in dataTable)
        {
          row++;
          col = 1;
          cells[row, col++] = item.Category;
          cells[row, col++] = item.IsPublic ? item.FieldName : "Reserved";
          cells[row, col++] = item.IsPublic ? item.Description : "";
          cells[row, col++] = item.Size;
        }
        app.Visible = true;
      }
    }

    private void InsertAbove(object sender, RoutedEventArgs e)
    {
      int selectedIndex = dataGrid.SelectedIndex;
      if (selectedIndex == -1) return;
      dataTable.Insert(selectedIndex, new Item());
    }

    private void InsertBelow(object sender, RoutedEventArgs e)
    {
      int selectedIndex = dataGrid.SelectedIndex;
      if (selectedIndex == -1) return;
      int indexToInsert = selectedIndex + 1;
      if (indexToInsert > dataTable.Count) return;
      dataTable.Insert(indexToInsert, new Item());
    }

    private void RemoveRow(object sender, RoutedEventArgs e)
    {
      int selectedIndex = dataGrid.SelectedIndex;
      if (selectedIndex == -1) return;
      if (selectedIndex >= dataTable.Count) return;
      dataTable.RemoveAt(selectedIndex);
      if (selectedIndex >= dataTable.Count) return;
      dataGrid.SelectedIndex = selectedIndex;
      dataGrid.Focus();
    }
  }

  public class Item
  {
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

    public string Category { get; set; }
    public string FieldName { get; set; }
    public string Description { get; set; }
    public int Size { get; set; }
    public bool IsPublic { get; set; }
    public string Comment { get; set; }
  }
}
