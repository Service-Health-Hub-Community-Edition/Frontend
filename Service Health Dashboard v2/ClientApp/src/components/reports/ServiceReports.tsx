import * as React from 'react';
import { ServiceIncidentReport } from './ServiceIncidentReport';

export class SharePointOnlineReport extends React.Component {

    render() {
        return <ServiceIncidentReport service="SharePoint Online" serviceDisplayName="SharePoint Online"  />
    };
}

export class ODBOnlineReport extends React.Component {

    render() {
        return <ServiceIncidentReport service="OneDrive for Business" serviceDisplayName="OneDrive for Business" />
    };
}

export class SFBOnlineReport extends React.Component {

    render() {
        return <ServiceIncidentReport service="Skype for Business" serviceDisplayName="Skype for Business" />
    };
}

export class ExchangeOnlineReport extends React.Component {

    render() {
        return <ServiceIncidentReport service="Exchange Online" serviceDisplayName="Exchange Online" />
    };
}

export class MicrosoftTeamsOnlineReport extends React.Component {

    render() {
        return <ServiceIncidentReport service="Microsoft Teams" serviceDisplayName="Microsoft Teams" />
    };
}

export class M365SuiteReport extends React.Component {

    render() {
        return <ServiceIncidentReport service="Microsoft 365 suite" serviceDisplayName="Microsoft 365 Suite" />
    };
}

export class Dynamics365Report extends React.Component {

    render() {
        return <ServiceIncidentReport service="Dynamics 365 Apps" serviceDisplayName="Dynamics 365 Apps" />
    };
}

export class PlannerReport extends React.Component {

    render() {
        return <ServiceIncidentReport service="Planner" serviceDisplayName="Planner" />
    };
}