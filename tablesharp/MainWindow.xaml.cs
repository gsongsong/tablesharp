using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace tablesharp
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

    private void BindData(ObservableCollection<Item> dataTable)
    {
      dataGrid.DataContext = dataTable;
      dataGrid.ItemsSource = dataTable;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
      dataTable = new ObservableCollection<Item>();
      BindData(dataTable);
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
        BindData(dataTable);
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
      var app = new Excel.Application();
      app.Workbooks.Add();
      Excel._Worksheet ws = app.ActiveSheet;
      Excel.Range cells = ws.Cells;
      Tuple<int, int> addr = new Tuple<int, int>(1, 1);
      addr = Item.FillHeader(cells, addr);
      foreach (Item item in dataTable)
      {
        addr = item.FillRow(cells, new Tuple<int, int>(addr.Item1 + 1, 1));
      }
      app.Visible = true;
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

    private void OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
    {
      Item.OnAutoGeneratingColumn(e);
    }
  }
}
