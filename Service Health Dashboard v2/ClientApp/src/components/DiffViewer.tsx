import * as React from 'react';
import { Text, DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from '@fluentui/react';

enum ItemDiffMode {
    Value,
    Object,
    Array
}

interface IItemDiff {
    name: string
    mode: ItemDiffMode
    oldValue: any
    newValue: any
}

export interface IDiffViewerState {
    itemDiff: IItemDiff[];
}

export class DiffViewer extends React.Component<{ diffObject: any }, IDiffViewerState> {
    constructor(props: { diffObject: any }) {
        super(props);

        this.state = {
            itemDiff: []
        };
    }

    public render() {
        const {
            itemDiff } = this.state;

        return (
            <div className='content'>
                {this._getObjectComponents(itemDiff, 0)}
           </div>
        );
    }

    componentDidMount() {
        var itemDiff: IItemDiff[] = this._processObject(this.props.diffObject.originalValues, this.props.diffObject.updatedValues);

        this.setState({
            itemDiff: itemDiff
        });
    }

    private _getObjectComponents(items: IItemDiff[], level: number) {
        return items.map((item) => item.mode === ItemDiffMode.Value ? (
            <div className='row' style={{ padding: '6px', paddingLeft: level * 16+'px' }} >
                <div className='col-3' style={{ overflow: 'hidden' }}>
                    <Text>{item.name}</Text>
                </div>
                <div className='col'>
                    <Text style={{
                        margin: '4px',
                        padding: '3px',
                        backgroundColor: 'rgba(223,246,221,1)'
                    }} >{item.newValue}</Text>

                    <Text style={{
                        margin: '4px',
                        padding: '3px',
                        backgroundColor: 'rgba(234,234,234,1)',
                        textDecoration: 'line-through'
                    }}>{item.oldValue}</Text>
                </div>
            </div>
        ) : item.mode === ItemDiffMode.Object ? (
                <div>
                    <div className='row' style={{ padding: '6px', paddingTop: '18px', paddingLeft: level * 16+'px' }} >
                        <div className='col' >
                            <Text><b>{item.name}</b></Text>
                        </div>
                    </div>
                    {this._getObjectComponents(item.newValue, level + 1)}
                </div>
        ) : item.mode === ItemDiffMode.Array ? (
                <div>
                    <div className='row' style={{ padding: '6px', paddingTop: '18px', paddingLeft: level * 16 + 'px' }} >
                        <div className='col'>
                            <Text><b>{item.name}</b></Text>
                        </div>
                    </div>
                    {this._getObjectArray(item.newValue, level)}
                </div>
        ): (<div />));
    }

    private _getObjectArray(objectArray: any, level: number) {
        if (objectArray) {
            var columns: IColumn[] = [];
            var columnNames: string[] = [];

            for (const obj of objectArray) {
                for (const key of Object.keys(obj))
                    if (!columnNames.find(cn => cn === key))
                        columnNames.push(key);
            }

            for (const columnName of columnNames)
                columns.push({
                    key: columnName,
                    name: columnName,
                    minWidth: 200,
                    isResizable: true,
                    onRender: (item: any) => {
                        return typeof item[columnName] === 'object' ?
                            (<Text>{JSON.stringify(item[columnName])}</Text >) : 
                            (<Text>{item[columnName]}</Text >)
                    }
                });

            return (<div className='row' style={{ paddingLeft: level * 16 + 'px' }} >
                <div className='col'>
                <DetailsList
                    items={objectArray}
                    compact={true}
                    columns={columns}
                    setKey='id'
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                    isHeaderVisible={true}
                        isSelectedOnFocus={true}
                        
                    />
                </div>
            </div>);
        } else
            return (<div />);
    }

    private _processObject(obj1: any, obj2: any): IItemDiff[] {
        var itemDiff: IItemDiff[] = [];
        var processedKeys: string[] = [];

        if (obj1)
            for (const key of Object.keys(obj1)) {
                if (Array.isArray(obj1[key])) {
                    itemDiff.push({
                        name: key,
                        mode: ItemDiffMode.Array,
                        oldValue: obj1[key],
                        newValue: obj2 ? obj2[key] : null
                    });
                } else if (typeof obj1[key] === 'object' && obj1[key] !== null) {
                    var childItemDiff = this._processObject(obj1[key], obj2 ? obj2[key] : null);
                    itemDiff.push({
                        name: key,
                        mode: ItemDiffMode.Object,
                        oldValue: null,
                        newValue: childItemDiff
                    });
                } else {
                    itemDiff.push({
                        name: key,
                        mode: ItemDiffMode.Value,
                        oldValue: obj1[key],
                        newValue: obj2 ? obj2[key] : null
                    });
                }
                processedKeys.push(key);
            };

        if (obj2) {
            for (const key of Object.keys(obj2)) {
                if (!processedKeys.find((k) => k === key)) {
                    if (Array.isArray(obj2[key])) {
                        itemDiff.push({
                            name: key,
                            mode: ItemDiffMode.Array,
                            oldValue: obj1[key],
                            newValue: obj2[key]
                        });
                    } else if (typeof obj2[key] === 'object' && obj2[key] !== null) {
                        var childItemDiff = this._processObject(obj1[key], obj2[key]);
                        itemDiff.push({
                            name: key,
                            mode: ItemDiffMode.Object,
                            oldValue: null,
                            newValue: childItemDiff
                        });
                    } else {

                        itemDiff.push({
                            name: key,
                            mode: ItemDiffMode.Value,
                            oldValue: null,
                            newValue: obj2[key]
                        });
                    }
                    processedKeys.push(key);
                }
            };
        } else {
            for (const key of processedKeys) {

                if (Array.isArray(obj1[key])) {
                    itemDiff.push({
                        name: key,
                        mode: ItemDiffMode.Array,
                        oldValue: obj1[key],
                        newValue: null
                    });
                } else if (typeof obj1[key] === 'object' && obj1[key] !== null) {
                    itemDiff.push({
                        name: key,
                        mode: ItemDiffMode.Object,
                        oldValue: null,
                        newValue: null
                    });
                } else {
                    itemDiff.push({
                        name: key,
                        mode: ItemDiffMode.Value,
                        oldValue: obj1[key],
                        newValue: null
                    });
                }
            }
        }

        return itemDiff;
    }
}