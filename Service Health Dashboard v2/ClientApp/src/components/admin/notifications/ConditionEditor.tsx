import * as React from 'react';
import { createRef, FC } from 'react';

import {
    Dropdown, IDropdownOption, ISelectableOption,
    TextField, Checkbox, Text, ITextProps, Label,
    ICommandBarProps, IObjectWithKey, SelectionMode,
    DialogFooter, PrimaryButton, DefaultButton
} from '@fluentui/react';
import { InfoTip } from '@m365-admin/info-tip';
import { IRoutingCondition, IRoutingElement, IComponent, IEntityProperty } from './Routing';

export const multiValueOperators = ['in', 'notin'];
export const operatorList: IDropdownOption[] = [
    { key: 'eq', text: '==' },
    { key: 'ne', text: '<>' },
    { key: 'gt', text: '>' },
    { key: 'ge', text: '>=' },
    { key: 'lt', text: '<' },
    { key: 'le', text: '<=' },
    { key: 'like', text: 'like' },
    { key: 'notlike', text: 'not like' },
    { key: 'match', text: 'match' },
    { key: 'notmatch', text: 'not match' },
    { key: 'contains', text: 'contains' },
    { key: 'notcontains', text: 'doesn\'t contain' },
    { key: 'in', text: 'in' },
    { key: 'notin', text: 'not in' }
];

interface IRoutingConditionState {
    originalCondition: IRoutingCondition;
    condition: IRoutingCondition;
    component?: IComponent;
    entityPropertyDropdown: IDropdownOption[];
    initialized: boolean;
}

var emptyCondition: IRoutingCondition = {
    key: '',
    property: '',
    operator: '',
    value: '',
    values: [],
    logicOperator: ''
}

export class ConditionEditor extends React.Component<{
    condition?: IRoutingCondition,
    component?: IComponent,
    onSave?: (condition: IRoutingCondition) => void,
    onCancel?: () => void
}, IRoutingConditionState> {  

    constructor(props: {
        condition?: IRoutingCondition,
        component?: IComponent,
        onSave?: (condition: IRoutingCondition) => void,
        onCancel?: () => void
    }) {
        super(props);

        var condition: IRoutingCondition = emptyCondition;

        if (this.props.condition)
            condition = Object.assign({} , this.props.condition);
        else
            condition = Object.assign({} , emptyCondition);

        this.state = {
            originalCondition: this.props.condition ? this.props.condition : emptyCondition,
            condition: condition,
            component: this.props.component,
            entityPropertyDropdown: [],
            initialized: false
        };
    }

    public render() {
        const {
            condition, initialized, entityPropertyDropdown
        } = this.state;

        if (!initialized)
            return (<div />);
        
        return (
            <div className="container" style={{ maxWidth: '97%' }}>
                <div className="row">
                    <div className="col">
                        <Dropdown
                            label="Property"
                            options={entityPropertyDropdown}
                            defaultSelectedKey={condition.property}
                            onChange={this._onChangeEntityPropertyDropdown}
                            accessKey='property'
                        />
                    </div>
                </div>
                    <div className="row">
                        <div className="col">
                            <Dropdown
                                label="Operator"
                                options={operatorList}
                                defaultSelectedKey={condition.operator ? condition.operator : 'eq' }
                                onChange={this._onChangeOperatorDropdown}
                                accessKey='operator'
                            />
                        </div>
                    </div>
                        {multiValueOperators.includes(condition.operator) ? (
                            <div className="row">
                            <div className="col" style={{paddingRight: '0px' }}>
                                    <TextField
                                        label="Values"
                                        multiline={true}
                                        defaultValue={condition.values ? condition.values.join('\n') : ''}
                                        onChange={this._onChangeValueField}
                                accessKey='values'
                                    />
                        </div>
                            <div className="col-auto" style={{ paddingLeft: '0px', marginTop: '18px' }} >
                                    <InfoTip
                                      info="For multiple values, enter one value per line."
                                      iconButtonProps={{ ariaLabel: 'More info' }}
                            />
                                </div>
                            </div>
                ) : (
                        <div className="row">
                            <div className="col">
                                    <TextField
                                        label="Value"
                                        multiline={false}
                                        defaultValue={condition.value ? condition.value : ''}
                                        onChange={this._onChangeValueField}
                                        accessKey='value'
                                />
                            </div>
                        </div>)}
                        
                    <div className="row">
                        <div className="col">
                            <Dropdown
                                label="Logic operator"
                                options={[
                                    { key: '', text: '' },
                                    { key: 'or', text: 'or' },
                                    { key: 'and', text: 'and' }
                                ]}
                                defaultSelectedKey={condition.logicOperator ? condition.logicOperator : ''}
                                onChange={this._onChangeLogicOperatorDropdown}
                                accessKey='logicOperator'
                            />
                        </div>
                    </div>

                <div className="row">
                    <div className="col">
                        <DialogFooter>
                            <PrimaryButton onClick={this._saveCondition} text="Save" />
                            <DefaultButton onClick={this._cancelEdit} text="Cancel" />
                        </DialogFooter>
                    </div>
                </div>
            </div>
        );
    }

    componentDidMount() {
        var entityPropertyDropdown: IDropdownOption[] = [];

        if (this.props.component !== undefined && this.props.component.entityProperties !== undefined)
            for (const entityProperty of this.props.component.entityProperties) {
                if (!entityProperty.hidden)
                    entityPropertyDropdown.push({
                        key: entityProperty.entityProperty,
                        text: entityProperty.displayName
                    });
            }

        entityPropertyDropdown = _copyAndSort(entityPropertyDropdown, 'text', false);

        this.setState({
            entityPropertyDropdown: entityPropertyDropdown,
            initialized: true
        });
    }

    private _onChangeValueField = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        var condition: IRoutingCondition = this.state.condition;
        var value: string = newValue ? newValue : '';
        switch (ev.currentTarget.accessKey) {
            case 'values':
                var valuesArray: string[] = value.split('\n');
                var conditionValues: string[] = [];
                for (const valueItem of valuesArray)
                    if (valueItem.trim() !== '')
                        conditionValues.push(valueItem);
                condition.values = conditionValues;
                break;
            default:
                condition.value = newValue ? newValue : '';
                break;
        }

        this.setState({
            condition: condition
        });
    };

    private _onChangeEntityPropertyDropdown = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
        var condition: IRoutingCondition = this.state.condition;
        condition.property = option ? option.key.toString() : '';
        this.setState({
            condition: condition
        });
    };

    private _onChangeOperatorDropdown = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
        var condition: IRoutingCondition = this.state.condition;
        condition.operator = option ? option.key.toString() : '';
        this.setState({
            condition: condition
        });
    };

    private _onChangeLogicOperatorDropdown = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
        var condition: IRoutingCondition = this.state.condition;
        condition.logicOperator = option ? option.key.toString() : '';
        this.setState({
            condition: condition
        });
    };

    private _doNothing = (): void => {

    }

    private _saveCondition = (): void => {
        this.props.onSave ? this.props.onSave(this.state.condition) : this._doNothing();
    }

    private _cancelEdit = (): void => {
        this.props.onCancel ? this.props.onCancel() : this._doNothing();
    }
}

function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
}
