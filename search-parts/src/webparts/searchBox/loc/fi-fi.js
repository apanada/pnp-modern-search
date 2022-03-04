define([], function () {
    return {
        General: {
            DynamicPropertyDefinition: "Hakukysely"
        },
        PropertyPane: {
            SearchBoxSettingsGroup: {
                GroupName: "Hakukentän asetukset",
                PlaceholderTextLabel: "Hakukentässä näkyvä teksti",
                SearchInNewPageLabel: "Lähetä haku uudelle sivulle",
                PageUrlLabel: "Sivun URL",
                UrlErrorMessage: "Syötä validi URL.",
                QueryPathBehaviorLabel: "Toimintatapa",
                QueryInputTransformationLabel: "Kyselyn muunnostemplaatti",
                UrlFragmentQueryPathBehavior: "URL osio",
                QueryStringQueryPathBehavior: "Hakukyselyn parametri",
                QueryStringParameterName: "Parametrin nimi",
                QueryParameterNotEmpty: "Syötä arvo parametrille."
            },
            SearchBoxQueryNlpSettingsGroup: {
                GroupName: "Search query enhancement",
                GroupDescription: "Use this service to enhance the search query by detecting user intents and get more relevant search keywords using Microsoft LUIS.",
                EnableNlpPropertyLabel: "Use Natural Language Processing service",
                ServiceUrlLabel: "Service URL",
                ServiceUrlDescription: "Notice: make sure the server allows cross-origin calls from this domain ('{0}') in CORS settings.",
                ServiceUrlErrorMessage: "Please specify a valid URL",                
                UrlNotResolvedErrorMessage: "URL '{0}' can't be resolved. Error: '{1}'."
            },
            AvailableConnectionsGroup: {
                GroupName: "Tarjolla olevat yhteydet",
                UseDynamicDataSourceLabel: "Käytä dynaamista sisältölähdettä oletussyötteenä",
                QueryKeywordsPropertyLabel: ""
            },
            QuerySuggestionsGroup: {
                GroupName: "Kyselyehdotukset",
                EnableQuerySuggestions: "Salli kyselyehdotukset",
                EditSuggestionProvidersLabel: "Konfiguroi kyselyehdotusten tarjoajat",
                SuggestionProvidersLabel: "Kyselyehdotusten tarjoajat",
                SuggestionProvidersDescription: "Salli tai estä kyselyehdotusten tarjoajia.",
                EnabledPropertyLabel: "Sallittu",
                ProviderNamePropertyLabel: "Nimi",
                ProviderDescriptionPropertyLabel: "Kuvaus",
                DefaultSuggestionGroupName: "Suositeltu",
                NumberOfSuggestionsToShow: "Kyselyehdotusten määrä per ryhmä"
            },
            InformationPage: {
                Extensibility: {
                    PanelHeader: "Konfiguroi käynnistyksessä ladattavat laajennuskirjastot mukautetuille kyselyehdotusten tarjoajille",
                    PanelDescription: "Lisää/poista mukautetun laajennuskirjastosi ID:t tässä. Voit määrittää näyttönimen ja päättää, ladataanko kirjasto käynnistyksessä. Vain mukautetut kyselyehdotusten tarjoajat ladataan tässä.",
                }
            },

        },
        SearchBox: {
            DefaultPlaceholder: "Syötä hakusanat...",
            SearchButtonLabel: "Suorita haku"
        }
    }
});