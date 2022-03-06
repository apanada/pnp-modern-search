import { FilterType } from "@pnp/modern-search-extensibility";

export enum BuiltinFilterTemplates {
    CheckBox = 'CheckboxFilterTemplate',
    DateRange = 'DateRangeFilterTemplate',
    ComboBox = 'ComboBoxFilterTemplate',
    DateInterval = 'DateIntervalFilterTemplate',
    DateTime = 'DateTimeFilterTemplate',
    TaxonomyPicker = 'TaxonomyPickerFilterTemplate',
    KnowledgeRepository = "KnowledgeRepositoryFilterTemplate"
}

/**
 * Filter types configuration
 */
export const BuiltinFilterTypes = {
    [BuiltinFilterTemplates.CheckBox]: FilterType.Refiner,
    [BuiltinFilterTemplates.DateInterval]: FilterType.Refiner,
    [BuiltinFilterTemplates.ComboBox]: FilterType.Refiner,
    [BuiltinFilterTemplates.DateRange]: FilterType.StaticFilter,
    [BuiltinFilterTemplates.DateTime]: FilterType.StaticFilter,
    [BuiltinFilterTemplates.TaxonomyPicker]: FilterType.StaticFilter,
    [BuiltinFilterTemplates.KnowledgeRepository]: FilterType.StaticFilter,    
};