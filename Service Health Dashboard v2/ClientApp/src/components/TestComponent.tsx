import * as React from 'react';
import { Text, DetailsList, Selection, SelectionMode, IObjectWithKey } from '@fluentui/react';
import { Separator } from '@fluentui/react';

export interface ITestComponentState {
    message: string;
    items: any;
    selection: any;
}

export class TestComponent extends React.Component<{ }, ITestComponentState> {
    private selection = new Selection(
        {
            onSelectionChanged: () => this.setState({ selection: this._getSelection() }),
        }
    );

    constructor(props: { }) {
        super(props);
        
        this.state = {
            message: "",
            items: [
                {
                    id: 1,
                    key: 1,
                    name: 'Item 1'
                },
                {
                    id: 2,
                    key: 2,
                    name: 'Item 2'
                }
            ],
            selection: undefined
        };
    }

    public render() {
        const {
            message, items, selection } = this.state;

        return (
            <DetailsList
                items={items}
                selectionMode={SelectionMode.single}
                selection={this.selection}
                />
        );
    }

    componentDidMount() {
        this.setState({
            message: "",
            items: [
                {
                    id: 1,
                    key: 1,
                    name: 'Item 1'
                },
                {
                    id: 2,
                    key: 2,
                    name: 'Item 2'
                }
            ],
            selection: undefined
        });
    }

    private _getSelection(): IObjectWithKey[] {
        return this.selection.getSelection();
    }
}