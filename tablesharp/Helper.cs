using System.Collections.Generic;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace tablesharp
{
  static class Helper
  {
    public static Binding BindingHelper(string key)
    {
      return new Binding(key)
      {
        Mode = BindingMode.TwoWay,
        UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
      };
    }

    public static DataGridTemplateColumn TemplateColumnHelper(string header, FrameworkElementFactory factory)
    {
      DataTemplate textTemplate = new DataTemplate
      {
        VisualTree = factory,
      };
      DataGridTemplateColumn column = new DataGridTemplateColumn
      {
        CellTemplate = textTemplate,
        Header = header,
      };
      return column;
    }

    public static DataGridTemplateColumn MultilineTextColumn(string header, string key)
    {
      Binding binding = BindingHelper(key);
      FrameworkElementFactory textFactory = new FrameworkElementFactory(typeof(TextBox));
      textFactory.SetBinding(TextBox.TextProperty, binding);
      textFactory.SetValue(TextBox.AcceptsReturnProperty, true);
      return TemplateColumnHelper(header, textFactory);
    }

    public static DataGridTemplateColumn CheckboxColumn(string header, string key)
    {
      Binding binding = BindingHelper(key);
      FrameworkElementFactory checkboxFactory = new FrameworkElementFactory(typeof(CheckBox));
      checkboxFactory.SetBinding(CheckBox.IsCheckedProperty, binding);
      return TemplateColumnHelper(header, checkboxFactory);
    }

    public static void OnAutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e, Dictionary<string, Property> itemTypes)
    {
      bool propertyFound = itemTypes.TryGetValue(e.PropertyName, out Property property);
      string header = propertyFound ? property.Header : e.PropertyName;
      PropertyInfo propertyInfo = typeof(Item).GetProperty(e.PropertyName);
      if (propertyInfo.PropertyType == typeof(bool))
      {
        e.Column = Helper.CheckboxColumn(header, e.PropertyName);
      }
      else if (property.InputType == InputType.Multiline)
      {
        e.Column = Helper.MultilineTextColumn(header, e.PropertyName);
      }
      else
      {
        e.Column.Header = header;
        Binding binding = (e.Column as DataGridTextColumn).Binding as Binding;
        binding.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
      }
      e.Column.CanUserSort = false;
    }
  }
}
