using System.Diagnostics.CodeAnalysis;

namespace PeoplePickerSearchApp
{
    public class OdataResponse
    {
        [SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "API property name")]
        public string? value { get; set; }
    }
}
