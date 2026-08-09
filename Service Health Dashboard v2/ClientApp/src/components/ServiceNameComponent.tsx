import * as React from 'react';
import { Text, FontIcon, IIconProps, mergeStyles, mergeStyleSets, ITheme } from '@fluentui/react';
import { GlobalState } from './GlobalState';

interface IServiceInfoMapping {
    name: string;
    icon: string;
    color: any;
}

export interface IServiceComponentState {
    name: string;
    additionalInfo?: string;
    serviceInfoMapping: IServiceInfoMapping;
    viewState: ServiceComponentViewState;
}

export enum ServiceComponentViewState {
    ShowAll = 0,
    IconOnly = 1
}

export class ServiceComponent extends React.Component<{ name: string, additionalInfo?: string, viewState?: ServiceComponentViewState, isPadded?: boolean }, IServiceComponentState> {
    static contextType = GlobalState;

    classNames: any = {};
    serviceMapping: IServiceInfoMapping[] = [];

    constructor(props: { name: string, additionalInfo?: string, viewState?: ServiceComponentViewState, isPadded?: boolean }) {
        super(props);

        const defaultMapping: IServiceInfoMapping = { name: '.default', icon: 'OfficeLogo', color: this.classNames.Office };

        this.state = {
            name: "",
            additionalInfo: undefined,
            serviceInfoMapping: defaultMapping,
            viewState: props.viewState ? props.viewState : ServiceComponentViewState.ShowAll
        }
    }

    render() {
        const { name, additionalInfo, serviceInfoMapping, viewState } = this.state;

        this._refreshClassNames();

        return (
            <div className={'svcComponent'} style={{ alignItems: 'center', display: 'inline-flex', margin: this.props.isPadded ? '0 4px' : '-4px 0 0' }}>
                <FontIcon iconName={serviceInfoMapping.icon} className={serviceInfoMapping.color} />
                <Text variant='small' style={{ margin: '-4px 4px 4px 0', paddingTop: '8px' }} hidden={viewState === ServiceComponentViewState.IconOnly}>{name}{additionalInfo ? " (" + additionalInfo + ")" : ""}</Text>
            </div>
        );
    }

    componentDidMount() {
        this._setComponentState();
    }

    componentDidUpdate(prevProps: any) {
        if (this.props.name != prevProps.name ||
            this.props.additionalInfo != prevProps.additionalInfo ||
            this.props.viewState != prevProps.viewState) {
            this._setComponentState();
        }
    }

    private _refreshClassNames(): void {
        const iconClass = mergeStyles({
            fontSize: 16,
            height: 16,
            width: 16,
            margin: '-4px 4px 4px 4px'
        });

        let globalState: any = this.context;
        var theme: ITheme = globalState.getTheme();

        this.classNames = mergeStyleSets({
            Azure: [{ color: '#007fff' }, iconClass],
            Access: [{ color: '#a4373a' }, iconClass],
            Edge: [{ color: '#7719aa' }, iconClass],
            Exchange: [{ color: '#0078d4' }, iconClass],
            Excel: [{ color: '#217346' }, iconClass],
            Forms: [{ color: '#077568' }, iconClass],
            Graph: [{ color: '#7719aa' }, iconClass],
            Office: [{ color: '#d83b01' }, iconClass],
            OneDrive: [{ color: '#0078d4' }, iconClass],
            OneNote: [{ color: '#7719aa' }, iconClass],
            Planner: [{ color: '#31752f' }, iconClass],
            Project: [{ color: '#31752f' }, iconClass],
            PowerApps: [{ color: '#742774' }, iconClass],
            PowerAutomate: [{ color: 'rgb(0, 102, 255)' }, iconClass],
            PowerPoint: [{ color: '#b7472a' }, iconClass],
            Publisher: [{ color: '#077568' }, iconClass],
            SharePoint: [{ color: '#0078d4' }, iconClass],
            Skype: [{ color: '#0078d4' }, iconClass],
            Sway: [{ color: '#008272' }, iconClass],
            Teams: [{ color: '#6264a7' }, iconClass],
            ToDo: [{ color: '#0078d4' }, iconClass],
            Visio: [{ color: '#3955a3' }, iconClass],
            Word: [{ color: '#2b579a' }, iconClass],
            Yammer: [{ color: '#106ebe' }, iconClass],
            PowerBI: [{ color: '#f2c80f' }, iconClass],
            Default: [{ color: theme.semanticColors.bodyText }, iconClass],
        });

        this.serviceMapping = [
            { name: 'Azure Information Protection', icon: 'Product', color: this.classNames.Default },
            { name: 'Azure Active Directory', icon: 'AADLogo', color: this.classNames.Default },
            { name: 'Bookings', icon: 'BookingsLogo', color: this.classNames.Default },
            { name: 'Dynamics 365', icon: 'Dynamics365Logo', color: this.classNames.Default },
            { name: 'Dynamics 365 Business Central', icon: 'Dynamics365Logo', color: this.classNames.Default },
            { name: 'Finance and Operations Apps', icon: 'Dynamics365Logo', color: this.classNames.Default },
            { name: 'Excel', icon: 'ExcelLogo', color: this.classNames.Excel },
            { name: 'Exchange', icon: 'ExchangeLogo', color: this.classNames.Exchange },
            { name: 'Exchange Online', icon: 'ExchangeLogo', color: this.classNames.Exchange },
            { name: 'Exchange Online Protection', icon: 'DefenderApp', color: this.classNames.Default },
            { name: 'Forms', icon: 'OfficeFormsLogo', color: this.classNames.Forms },
            { name: 'Microsoft Forms', icon: 'OfficeFormsLogo', color: this.classNames.Forms },
            { name: 'Microsoft Defender ATP', icon: 'ShieldSolid', color: this.classNames.Default },
            { name: 'Microsoft Defender for Identity', icon: 'ShieldSolid', color: this.classNames.Default },
            { name: 'Microsoft Defender for Office 365', icon: 'ShieldSolid', color: this.classNames.Default },
            { name: 'Microsoft Edge', icon: 'EdgeLogo', color: this.classNames.Edge },
            { name: 'Microsoft Flow', icon: 'MicrosoftFlowLogo', color: this.classNames.PowerAutomate },
            { name: 'Microsoft Flow in Microsoft 365', icon: 'MicrosoftFlowLogo', color: this.classNames.PowerAutomate },
            { name: 'Microsoft Graph', icon: 'WindowsLogo', color: this.classNames.Graph },
            { name: 'Microsoft Power Automate', icon: 'MicrosoftFlowLogo', color: this.classNames.PowerAutomate },
            { name: 'Microsoft Power Automate in Microsoft 365', icon: 'MicrosoftFlowLogo', color: this.classNames.PowerAutomate },
            { name: 'Microsoft Project', icon: 'ProjectLogo32', color: this.classNames.Project },
            { name: 'Microsoft Teams', icon: 'TeamsLogo', color: this.classNames.Teams },
            { name: 'Microsoft To Do', icon: 'ToDoLogoInverse', color: this.classNames.ToDo },
            { name: 'Microsoft To-Do', icon: 'ToDoLogoInverse', color: this.classNames.ToDo },
            { name: 'Mobile Device Management for Office 365', icon: 'Product', color: this.classNames.Default },
            { name: 'Microsoft Intune', icon: 'Product', color: this.classNames.Default },
            { name: 'Microsoft Kaizala', icon: 'KaizalaLogo', color: this.classNames.Default },
            { name: 'Skype for Business', icon: 'SkypeForBusinessLogo', color: this.classNames.Skype },
            { name: 'PowerApps', icon: 'PowerAppsLogo', color: this.classNames.PowerApps },
            { name: 'Power Apps', icon: 'PowerAppsLogo', color: this.classNames.PowerApps },
            { name: 'Power Apps in Microsoft 365', icon: 'PowerAppsLogo', color: this.classNames.PowerApps },
            { name: 'Power Automate', icon: 'PowerAppsLogo', color: this.classNames.PowerAutomate },
            { name: 'Power BI', icon: 'PowerBILogo', color: this.classNames.PowerBI },
            { name: 'PowerPoint', icon: 'PowerPointLogo', color: this.classNames.PowerPoint },
            { name: 'Office Online', icon: 'OfficeLogo', color: this.classNames.Office },
            { name: 'Office Client Applications', icon: 'OfficeLogo', color: this.classNames.Office },
            { name: 'OneDrive', icon: 'OneDriveLogo', color: this.classNames.OneDrive },
            { name: 'OneDrive for Business', icon: 'OneDriveLogo', color: this.classNames.OneDrive },
            { name: 'OneNote', icon: 'OneNoteLogo', color: this.classNames.OneNote },
            { name: 'Identity Service', icon: 'Product', color: this.classNames.Default },
            { name: 'Microsoft 365 suite', icon: 'AdminALogo32', color: this.classNames.Default },
            { name: 'Microsoft 365 Admin Center', icon: 'AdminALogo32', color: this.classNames.Default },
            { name: 'Microsoft Cloud App Security', icon: 'DefenderApp', color: this.classNames.Default },
            { name: 'Office Subscription', icon: 'OfficeLogo', color: this.classNames.Office },
            { name: 'Outlook', icon: 'OutlookLogo', color: this.classNames.Exchange },
            { name: 'Outlook.com', icon: 'OutlookLogo', color: this.classNames.Exchange },
            { name: 'Outlook Mobile', icon: 'OutlookLogo', color: this.classNames.Exchange },
            { name: 'Outlook for the Web', icon: 'OutlookLogo', color: this.classNames.Exchange },
            { name: 'Planner', icon: 'PlannerLogo', color: this.classNames.Planner },
            { name: 'Project for the web', icon: 'ProjectLogo32', color: this.classNames.Project },
            { name: 'Project Online', icon: 'ProjectLogo32', color: this.classNames.Project },
            { name: 'SharePoint', icon: 'SharepointLogo', color: this.classNames.SharePoint },
            { name: 'SharePoint Online', icon: 'SharepointLogo', color: this.classNames.SharePoint },
            { name: 'SharePoint Syntex', icon: 'SharepointLogo', color: this.classNames.SharePoint },
            { name: 'Microsoft Stream', icon: 'StreamLogo', color: this.classNames.Default },
            { name: 'Sway', icon: 'SwayLogo32', color: this.classNames.Sway },
            { name: 'Yammer', icon: 'YammerLogo', color: this.classNames.Yammer },
            { name: 'Yammer Enterprise', icon: 'YammerLogo', color: this.classNames.Yammer },
            { name: 'Visio', icon: 'VisioLogo', color: this.classNames.Visio },
            { name: 'Whiteboard', icon: 'WhiteBoardApp16', color: this.classNames.Default },
            { name: 'Windows', icon: 'WindowsLogo', color: this.classNames.Default },
            { name: 'Windows 365', icon: 'WindowsLogo', color: this.classNames.Default },
            { name: 'Word', icon: 'WordLogo', color: this.classNames.Word },
            { name: '.default', icon: 'OfficeLogo', color: this.classNames.Office }
        ];
    }

    private _setComponentState(): void {
        var mappingMatch: IServiceInfoMapping | undefined = this.serviceMapping.find(
            m => m.name.trim().toLowerCase() === this.props.name.trim().toLowerCase());

        const defaultMapping: IServiceInfoMapping = { name: '.default', icon: 'OfficeLogo', color: this.classNames.Office };

        this.setState({
            name: this.props.name,
            additionalInfo: this.props.additionalInfo,
            serviceInfoMapping: mappingMatch ? mappingMatch : defaultMapping,
            viewState: this.props.viewState ? this.props.viewState : ServiceComponentViewState.ShowAll
        });
    }

}