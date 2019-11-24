using System;
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
    private readonly string member = "public {0} {1} {{ get; set; }}";
    private readonly string dictItem = "{{ \"{0}\", new Property(\"{1}\") }},";
    private readonly string dicItemMultiline = "{{ \"{0}\", new Property(\"{1}\", InputType.Multiline) }},";
    private readonly string consructorArgs = "{0} {1}";
    private readonly string constructor = "{0} = {1};";
    private readonly string cell = "cells[row, col++] = \"{0}\";";

    static void Main(string[] args)
    {
      Assembly assembly = Assembly.GetExecutingAssembly();
      Stream stream = assembly.GetManifestResourceStream("itembuilder.Item.json");
      StreamReader streamReader = new StreamReader(stream);
      string json = streamReader.ReadToEnd();
      Console.WriteLine(json);
      Item[] items = JsonSerializer.Deserialize<Item[]>(json);
    }
  }
}
