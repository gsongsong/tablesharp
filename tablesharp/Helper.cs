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
  }
}
