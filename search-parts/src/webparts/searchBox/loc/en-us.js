define([], function() {
    return {
        General: {
            DynamicPropertyDefinition: "Search query"
        },
        PropertyPane: {
            SearchBoxSettingsGroup: {
                GroupName: "Search box settings",
                PlaceholderTextLabel: "Placeholder text to display in the search box",
                SearchInNewPageLabel: "Send the query to a new page",
                PageUrlLabel: "Page URL",
                UrlErrorMessage: "Please provide a valid URL.",
                QueryPathBehaviorLabel: "Method",
                QueryInputTransformationLabel: "Query input transformation template",
                UrlFragmentQueryPathBehavior: "URL fragment",
                QueryStringQueryPathBehavior: "Query string parameter",
                QueryStringParameterName: "Parameter name",
                QueryParameterNotEmpty: "Please provide a value for the parameter."
            },
            SearchBoxQueryNlpSettingsGroup: {
                GroupName: "Search query enhancement",
                GroupDescription: "Use this service to enhance the search query by detecting user intents and get more relevant search keywords using Microsoft LUIS.",
                EnableNlpPropertyLabel: "Use Natural Language Processing service",
                ServiceUrlLabel: "Service URL",
                ServiceUrlDescription: "Notice: make sure the server allows cross-origin calls from this domain ('{0}') in CORS settings.",
                ServiceUrlErrorMessage: "Please specify a valid URL",
                UseStagingEndpoint: "Use LUIS staging endpoint",
                UrlNotResolvedErrorMessage: "URL '{0}' can't be resolved. Error: '{1}'."
            },
            AvailableConnectionsGroup: {
                GroupName: "Available connections",
                UseDynamicDataSourceLabel: "Use dynamic data source as default input",
                QueryKeywordsPropertyLabel: ""
            },
            QuerySuggestionsGroup: {
                GroupName: "Query suggestions",
                EnableQuerySuggestions: "Enable query suggestions",
                EditSuggestionProvidersLabel: "Configure available providers",
                SuggestionProvidersLabel: "Suggestion providers",
                SuggestionProvidersDescription: "Enable or disable individual suggestion providers.",
                EnabledPropertyLabel: "Enabled",
                ProviderNamePropertyLabel: "Name",
                ProviderDescriptionPropertyLabel: "Description",
                DefaultSuggestionGroupName: "Recommended",
                NumberOfSuggestionsToShow: "Number of suggestions to show per group"
            },
            InformationPage: {
                Extensibility: {
                    PanelHeader: "Configure extensibility libraries to load at startup for custom suggestions providers",
                    PanelDescription: "Add/Remove your custom extensibility library IDs here. You can specify a display name and decide if the library should be loaded or not at startup. Only custom suggestions providers will be loaded here.",
                }
            },
            
        },
        SearchBox: {
            DefaultPlaceholder: "Enter your search terms...",
            SearchButtonLabel: "Search"
        }
    }
});