using System.Diagnostics.CodeAnalysis;

namespace PeoplePickerSearchApp
{
    public class PeoplePickerSearchUserPayload
    {
        [SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "API property name")]
        public PeoplePickerSearchUserQueryParams? queryParams { get; init; }
    }

    public class PeoplePickerSearchUserQueryParams
    {
        public int PrincipalType { get; init; }

        public int PrincipalSource { get; init; }

        public string? QueryString { get; init; }

        public bool AllowMultipleEntities { get; init; }

        public int MaximumEntitySuggestions { get; init; }

        public bool UseSubstrateSearch { get; init; }
    }
}
