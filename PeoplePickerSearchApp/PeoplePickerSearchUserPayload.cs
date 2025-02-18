namespace PeoplePickerSearchApp
{
    public class PeoplePickerSearchUserPayload
    {
        public PeoplePickerSearchUserQueryParams? queryParams { get; init; }
    }

    public class PeoplePickerSearchUserQueryParams
    {
        //public RestApiTypeData __metadata { get; } = new RestApiTypeData {
        //    type = "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters",
        //};

        public int PrincipalType { get; init; }

        public int PrincipalSource { get; init; }

        public string? QueryString { get; init; }

        public bool AllowMultipleEntities { get; init; }

        public int MaximumEntitySuggestions { get; init; }

        public bool UseSubstrateSearch { get; init; }
    }

    public class RestApiTypeData
    {
        public string? type { get; init; }
    }
}
