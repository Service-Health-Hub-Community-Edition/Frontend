import { IContextualMenuItem } from '@fluentui/react';

export enum FilterOperator {
    eq,
    gt,
    ge,
    lt,
    le,
    contains
}

export interface IFilterProps {
    filter?: IContextualMenuItem[] | undefined;
    property: string | string[];
    operator?: FilterOperator;
    value: string | string[] | boolean | Date | undefined;
    convertValue?(value: string | string[] | boolean | Date | undefined, operator?: FilterOperator): string | string[] | boolean | boolean[] | Date | Date[] | undefined;
}

export interface IFilter {
    [key: string]: IFilterProps;
}


export interface IFilterCollection {
    [key: string]: IFilter;
}

export class Filter {
    public _filter: IFilter;
    private _onFilterCleared?: () => void;
    private _onFilterUpdated?: () => void;

    constructor(filter: IFilter, onFilterCleared?: () => void, onFilterUpdated?: () => void) {
        this._filter = filter;
        this._onFilterCleared = onFilterCleared;
        this._onFilterUpdated = onFilterUpdated;
    }

    public setFilter(filter: IFilter) {
        this._filter = filter;
    }

    private _textMatch(property: string | string[], item: any, searchString: string | undefined): boolean {
        if (!searchString)
            return true;

        var match: boolean = false;

        const itemValueMatch = (itemValue: any | undefined, searchString: string): boolean => {
            return Array.isArray(itemValue) ?
                itemValue?.find(v => v.toLowerCase().indexOf(searchString.toLowerCase) > -1) !== undefined
                :
                (itemValue !== null ? itemValue?.toLowerCase().indexOf(searchString.toLowerCase()) > -1 : false);
        }

        if (Array.isArray(property)) {
            for (const p of property) {
                let itemValue = item[p];
                match = match || itemValueMatch(itemValue, searchString);
            }
        } else {
            let itemValue = item[property];
            match = match || itemValueMatch(itemValue, searchString);
        }

        return match;
    }

    public filterItems(items: any[]): any[] {
        var filteredItems: any[] = Object.assign(items);

        for (const key of Object.keys(this._filter)) {
            var filter = this._filter[key as keyof IFilter];
            var value = filter.convertValue ? filter.convertValue(filter.value, filter.operator) : filter.value;

            if (filter.filter === undefined)
                filteredItems = value !== "" ? filteredItems.filter(i => this._textMatch(filter.property, i, value as string)) : filteredItems;
            else {
                if (value !== undefined) {
                    var filterArray: any[] = Array.isArray(value) ? value : value !== undefined ? [value] : [];
                    var filterProperties: any[] = Array.isArray(filter.property) ? filter.property : [filter.property];

                    if (filterArray.length > 0) {
                        switch (filter.operator) {
                            case FilterOperator.eq:
                                filteredItems = filteredItems.filter(i => filterProperties.some(p => filterArray.some(filterElement => {
                                    let v = i[p];
                                    return v == filterElement
                                })));
                                break;

                            case FilterOperator.lt:
                                filteredItems = filteredItems.filter(i => filterProperties.some(p => filterArray.some(filterElement => {
                                    let v = i[p];
                                    return v <= filterElement
                                })));
                                break;

                            case FilterOperator.le:
                                filteredItems = filteredItems.filter(i => filterProperties.some(p => filterArray.some(filterElement => {
                                    let v = i[p];
                                    return v <= filterElement
                                })));
                                break;

                            case FilterOperator.gt:
                                filteredItems = filteredItems.filter(i => filterProperties.some(p => filterArray.some(filterElement => {
                                    let v = i[p];
                                    return v > filterElement
                                })));
                                break;

                            case FilterOperator.ge:
                                filteredItems = filteredItems.filter(i => filterProperties.some(p => filterArray.some(filterElement => {
                                    let v = i[p];
                                    return v >= filterElement
                                })));
                                break;
                            default:
                                filteredItems = filteredItems.filter(i => filterProperties.some(p => filterArray.some(filterElement => {
                                    let v = i[p];
                                    return Array.isArray(v) ? v.indexOf(filterElement) >= 0 : v == filterElement;
                                })));
                        }
                    }
                }
            }
        }

        return filteredItems
    }

    public setFilterSearchItem(filterKey: string, value?: string): void {
        var filterObject: IFilterProps | undefined = filterKey in this._filter ? this._filter[filterKey] : undefined;

        if (filterObject && filterObject.filter === undefined) {
            filterObject.value = value;

            if (this._onFilterUpdated)
                this._onFilterUpdated();
        }
    }


    public setFilterItem(filterKey: string, key: string, singleItemSelection?: boolean): void {
        var filterObject: IFilterProps | undefined = filterKey in this._filter ? this._filter[filterKey] : undefined;

        if (filterObject) {
            var filterValue: string[] = Array.isArray(filterObject.value) ? filterObject.value : [];
            var changed: boolean = false;
            var option: IContextualMenuItem | undefined = filterObject.filter?.find((o) => o.key === key);

            if (option !== undefined) {
                option.checked = !option.checked;

                if (singleItemSelection) {
                    if (option.checked) {
                        filterValue = [];
                        filterValue.push(option.text!);
                        changed = true;
                    } else {
                        filterValue = [];
                        changed = true;
                    }

                    if (filterObject.filter)
                        for (const filterOption of filterObject.filter)
                            if (filterOption.key !== key)
                                filterOption.checked = false;

                } else {
                    if (option.checked) {
                        if (filterValue.filter(s => s.toLowerCase() === option!.text!.toLowerCase()).length === 0) {
                            filterValue.push(option.text!);
                            changed = true;
                        }
                    } else {
                        var idx: number = filterValue.indexOf(option!.text!);
                        if (idx > -1) {
                            filterValue.splice(idx, 1);
                            changed = true;
                        }
                    }
                }
            }

            filterObject.value = filterValue;

            if (changed) {
                this._filter[filterKey] = filterObject;
                if (this._onFilterUpdated)
                    this._onFilterUpdated();
            }
        }
    }

    private resetFilterOptions(options: IContextualMenuItem[]): void {
        for (const option of options)
            option.checked = false;
    }

    public clearFilter(): void {
        for (const key of Object.keys(this._filter)) {
            var filterObject = this._filter[key as keyof IFilter];
            if (filterObject.filter !== undefined) {
                this.resetFilterOptions(filterObject.filter);
                filterObject.value = [];
            } else {
                filterObject.value = undefined;
            }
        }

        if (this._onFilterCleared)
            this._onFilterCleared();
    }
}