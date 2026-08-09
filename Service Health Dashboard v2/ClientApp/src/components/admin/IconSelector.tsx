import * as React from 'react';
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownProps, ISelectableOption } from '@fluentui/react';
import { Icon } from '@fluentui/react';
import { Label } from '@fluentui/react';
import { IStackTokens, Stack } from '@fluentui/react';
import { IconButton } from '@fluentui/react';

interface IIconSelectorState {
    selectedIcon: string | number;
}

const stackTokens: IStackTokens = { childrenGap: 20 };
const iconStyles = { marginRight: '8px' };

const dropdownStyles = { dropdown: { width: 300 } };

export class IconSelector extends React.Component<{
    selectedIcon?: string,
    onChange?: (icon: string) => void
}, IIconSelectorState> {
    dropdownOptions: IDropdownOption[] = [
        { key: 'Header', text: 'Services', itemType: DropdownMenuItemType.Header },
        { key: 'Dynamics365Logo', text: 'Dynamics 365', data: { icon: 'Dynamics365Logo' } },
        { key: 'ExcelLogo', text: 'Excel', data: { icon: 'ExcelLogo' } },
        { key: 'ExchangeLogo', text: 'Exchange', data: { icon: 'ExchangeLogo' } },
        { key: 'OfficeLogo', text: 'Office', data: { icon: 'OfficeLogo' } },
        { key: 'OneNoteLogo', text: 'OneNote', data: { icon: 'OneNoteLogo' } },
        { key: 'OutlookLogo', text: 'Outlook', data: { icon: 'OutlookLogo' } },
        { key: 'PowerApps', text: 'Power Apps', data: { icon: 'PowerApps' } },
        { key: 'PowerPointLogo', text: 'PowerPoint', data: { icon: 'PowerPointLogo' } },
        { key: 'SharepointLogo', text: 'SharePoint', data: { icon: 'SharepointLogo' } },
        { key: 'TeamsLogo', text: 'Teams', data: { icon: 'TeamsLogo' } },
        { key: 'divider_2', text: '-', itemType: DropdownMenuItemType.Divider },
        { key: 'Header2', text: 'Other', itemType: DropdownMenuItemType.Header },
        { key: 'CalculatorAddition', text: 'Add', data: { icon: 'CalculatorAddition' } },
        { key: 'AlertSolid', text: 'Alert', data: { icon: 'AlertSolid' } },
        { key: 'SingleBookmark', text: 'Bookmark', data: { icon: 'SingleBookmark' } },
        { key: 'Broom', text: 'Clean up', data: { icon: 'Broom' } },
        { key: 'Completed', text: 'Completed', data: { icon: 'Completed' } },
        { key: 'Dataflows', text: 'Data flow', data: { icon: 'Dataflows' } },
        { key: 'Calories', text: 'Fire', data: { icon: 'Calories' } },
        { key: 'Health', text: 'Health', data: { icon: 'Health' } },
        { key: 'IncidentTriangle', text: 'Incident', data: { icon: 'IncidentTriangle' } },
        { key: 'News', text: 'News', data: { icon: 'News' } },
        { key: 'Permissions', text: 'Permissions', data: { icon: 'Permissions' } },
        { key: 'CalculatorSubtract', text: 'Remove', data: { icon: 'CalculatorSubtract' } },       
        { key: 'ReportAdd', text: 'Report', data: { icon: 'ReportAdd' } },
        { key: 'Send', text: 'Send', data: { icon: 'Send' } },
        { key: 'Shield', text: 'Shield', data: { icon: 'Shield' } },
        { key: 'ShieldAlert', text: 'Shield (alert)', data: { icon: 'ShieldAlert' } },
        { key: 'SyncStatus', text: 'Sync', data: { icon: 'SyncStatus' } },
        { key: 'TaskSolid', text: 'Task', data: { icon: 'TaskSolid' } },
        { key: 'HighlightMappedShapes', text: 'Workflow', data: { icon: 'HighlightMappedShapes' } },
    ];

    constructor(props: { selectedIcon?: string}) {
        super(props);

        this.state = {
            selectedIcon: this.props.selectedIcon ? this.props.selectedIcon : ''
        };
    }

    public render() {
        const {
            selectedIcon } = this.state;

        return (
            <Dropdown
                placeholder="Select an icon"
                label="Icon"
                ariaLabel="Icon selection"
                onRenderPlaceholder={this._onRenderPlaceholder}
                onRenderTitle={this._onRenderTitle}
                onRenderOption={this._onRenderOption}              
                options={this.dropdownOptions}
                onChange={this._onIconSelectionChange}
                defaultSelectedKey={selectedIcon !== undefined && selectedIcon !== null ? selectedIcon : ''}
            />
        );
    }

    componentDidMount() {
        
    }

    _onIconSelectionChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
        this.props.onChange ? this.props.onChange(option ? option.key.toString() : '') : this._doNothing();

        this.setState({
            selectedIcon: option ? option.key : ''
        });
    }

    private _doNothing = (): void => {

    }

    private _onRenderPlaceholder = (props: IDropdownProps | undefined): JSX.Element => {
        if (props !== undefined)
            return (
                <div className="dropdown-placeholder">
                    <span>{props.placeholder}</span>
                </div>
            );
        else
            return (<div />);
    };

    private _onRenderLabel = (props: IDropdownProps | undefined): JSX.Element => {
        if (props !== undefined)
            return (
                <Stack horizontal verticalAlign="center">
                    <Label>{props.label}</Label>
                    <IconButton
                        iconProps={{ iconName: 'Info' }}
                        title="Info"
                        ariaLabel="Info"
                        styles={{ root: { marginBottom: -3 } }}
                    />
                </Stack>
            );
        else
            return (<div />);
    };

    private _onRenderOption = (option: ISelectableOption | undefined): JSX.Element => {
        if (option !== undefined)
            return (
                <div>
                    {option.data && option.data.icon && (
                        <Icon style={iconStyles} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
                    )}
                    <span>{option.text}</span>
                </div>
            );
        else
            return (<div />);
    };

    private _onRenderTitle = (options: ISelectableOption[] | undefined): JSX.Element => {

        if (options !== undefined && options.length > 0) {
            const option = options[0];

            return (
                <div>
                    {option.data && option.data.icon && (
                        <Icon style={iconStyles} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
                    )}
                    <span>{option.text}</span>
                </div>
            );
        } else
            return (<div/>);
    };
}