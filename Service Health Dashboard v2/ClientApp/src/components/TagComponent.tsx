import * as React from 'react';
import { Text, FontIcon, IIconProps, mergeStyles, mergeStyleSets, ITheme } from '@fluentui/react';
import { GlobalState } from './GlobalState';

interface ITagMapping {
    name: string;
    icon: string;
    color: any;
}

const iconClass = mergeStyles({
    fontSize: 16,
    height: 16,
    width: 16,
    margin: '-4px 0px 4px 4px'
});

const classNames = mergeStyleSets({
    New: [{ color: '#0078d4' }, iconClass],
    Retirement: [{ color: '#d83b01' }, iconClass],
    Updated: [{ color: '#31752f' }, iconClass],
    Warning: [{ color: '#f2c80f' }, iconClass],
    Admin: [{ color: '#007fff' }, iconClass],
    AdminImpact: [{ color: '#d83b01' }, iconClass],
    Default: [{ color: 'black' }, iconClass],
});

const serviceMapping: ITagMapping[] = [
    { name: 'New Feature', icon: 'AddIn', color: classNames.New },
    { name: 'Retirement', icon: 'PageRemove', color: classNames.Retirement },
    { name: 'Updated message', icon: 'Message', color: classNames.Updated },
    { name: 'User impact', icon: 'UserWarning', color: classNames.Warning },
    { name: 'Feature update', icon: 'ProductRelease', color: classNames.Updated },
    { name: 'Admin impact', icon: 'Admin', color: classNames.AdminImpact }
];

interface ITagState {
    name: string;
    compact: boolean;
    viewState: TagViewState;
    tagMapping: ITagMapping;
    iconTitle?: string;
}

export enum TagViewState {
    ShowAll = 0,
    IconOnly = 1,
    TextOnly = 2
}

export class Tag extends React.Component<{ name: string, compact?: boolean, viewState?: TagViewState, iconTitle?: string }, ITagState> {
    static contextType = GlobalState;

    constructor(props: { name: string, compact?: boolean, viewState?: TagViewState, iconTitle?: string }) {
        super(props);
        
        this.state = {
            name: "",
            compact: true,
            viewState: TagViewState.ShowAll,
            iconTitle: undefined,
            tagMapping: { name: "", icon: "", color: classNames.Default }
        }
    }

    render() {
        const { name, compact, tagMapping, viewState, iconTitle } = this.state;

        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();
        
        var margins: string = compact === true ? '0 4px' : '4px 4px 4px';
        var border: string = viewState === TagViewState.IconOnly ? "" : "1px solid " + theme.semanticColors.bodyDivider;
        var background: string = viewState === TagViewState.IconOnly ? "" : theme.semanticColors.menuBackground;

        if (name === undefined)
            return (<div />);
        else
            return (
                <div className={'tagComponent'} style={{ border: border, alignItems: 'center', display: 'inline-flex', margin: margins, background: background }}>
                    <FontIcon
                        iconName={tagMapping.icon}
                        className={tagMapping.color}
                        hidden={tagMapping.icon === '' || !(viewState === TagViewState.ShowAll || viewState === TagViewState.IconOnly)}
                        title={ iconTitle }
                    />
                    <Text
                        variant='xSmall'
                        style={{ margin: '-4px 4px 4px 4px', paddingTop: '8px' }}
                        hidden={name === null ? true : name.toLowerCase().trim() === '' || !(viewState === TagViewState.ShowAll || viewState === TagViewState.TextOnly)}
                    >
                        {name === null ? "" : (name.toLowerCase().trim() === '' || !(viewState === TagViewState.ShowAll || viewState === TagViewState.TextOnly) ? "" : name.toUpperCase())}
                    </Text>
                </div>
            );
    }

    componentDidMount() {
        this._setComponentState();
    }

    componentDidUpdate(prevProps: any) {
        if (this.props.name != prevProps.name ||
            this.props.compact != prevProps.compact ||
            this.props.viewState != prevProps.viewState ||
            this.props.iconTitle != prevProps.iconTitle) {

            this._setComponentState();
        }
    }

    private _setComponentState(): void {
        var mappingMatch: ITagMapping | undefined = this.props.name === undefined || this.props.name === null ? undefined : serviceMapping.find(
            m => m.name.trim().toLowerCase() === this.props.name.trim().toLowerCase());

        const defaultMapping: ITagMapping = { name: '', icon: '', color: classNames.Default };

        this.setState({
            name: this.props.name === null ? "" : this.props.name,
            compact: this.props.compact !== undefined ? this.props.compact : true,
            viewState: this.props.viewState !== undefined ? this.props.viewState : TagViewState.ShowAll,
            iconTitle: this.props.iconTitle,
            tagMapping: mappingMatch ? mappingMatch : defaultMapping
        });
    }
}