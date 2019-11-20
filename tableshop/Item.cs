namespace tableshop
{
  class Item
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
