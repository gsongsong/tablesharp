using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;

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

    private void InsertAbove(object sender, RoutedEventArgs e)
    {
      int selectedIndex = dataGrid.SelectedIndex;
      if (selectedIndex == -1) return;
      dataTable.Insert(selectedIndex, new Item());
    }

    private void SaveAs(object sender, RoutedEventArgs e)
    {
      SaveFileDialog saveFileDialog = new SaveFileDialog
      {
        DefaultExt = "json",
        Filter = "JSON files (.json)|*.json",
      };
      bool? dialogResult = saveFileDialog.ShowDialog();
      if (dialogResult == true)
      {
        string filename = saveFileDialog.FileName;
        string json = JsonSerializer.Serialize(dataTable);
        File.WriteAllText(filename, json);
      }
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
