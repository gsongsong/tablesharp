namespace tablesharp
{
  enum InputType { Default, Multiline }

  class Property
  {
    public string Header { get; }
    public InputType InputType { get; }

    public Property(string header, InputType inputType)
    {
      Header = header;
      InputType = inputType;
    }

    public Property(string header)
    {
      Header = header;
      InputType = InputType.Default;
    }
  }
}
