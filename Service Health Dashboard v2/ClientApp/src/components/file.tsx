import * as React from 'react';
import { IncidentDetails } from './IncidentDetails';
import { ServiceComponent } from './ServiceNameComponent';
import { Tag } from './TagComponent';

export class TestComponent extends React.Component {

    render() {
        return (
            <div>
                <div>
                    <ServiceComponent name='Office Online' />
                    <ServiceComponent name='SharePoint Online' />
                    <ServiceComponent name='Microsoft Teams' />
                    <ServiceComponent name='Microsoft Power Automate in Microsoft 365' />
                    <ServiceComponent name='Power Apps' />
                </div>

                <div>
                    <Tag name='Admin impact' />
                    <Tag name='Feature update' />
                    <Tag name='New feature' />
                    <Tag name='Retirement' />
                    <Tag name='Updated message' />
                    <Tag name='User impact' />
                </div>
            </div>
        );
    };
}