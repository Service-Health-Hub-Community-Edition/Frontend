import * as React from 'react';
import { Routing } from './Routing';
import { AdminLayout } from '../AdminLayout';

interface INotificationsRoutingState {
    entityId?: string
}

export class NotificationsRouting extends React.Component<{}, INotificationsRoutingState> {  

    constructor(props: {}) {
        super(props);

        const queryParams = new URLSearchParams(window.location.search);
        const entityId = queryParams.get('entityId');

        this.state = {
            entityId: entityId ? entityId : undefined
        };

    }

    public render() {
        const { entityId } = this.state;

        if (entityId === undefined)
            return (<div />);

        return (
            <AdminLayout>
                <Routing componentId={entityId} routeType='notificationManager' />
            </AdminLayout>
        );
    }

    componentDidMount() {

    }   
}