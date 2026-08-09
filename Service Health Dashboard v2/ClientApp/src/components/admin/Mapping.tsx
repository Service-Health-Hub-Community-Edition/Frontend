import * as React from 'react';
import { Text, Pivot, PivotItem } from '@fluentui/react';
import { AdminLayout } from "./AdminLayout";
import { ConfigEditor } from "./jsonConfigEditor/ConfigEditor";

interface IAdminMappingState {
    component: string
}

export class AdminMapping extends React.Component<{}, IAdminMappingState> {
    constructor(props: {}) {
        super(props);

        this.state = {
            component: "ServiceUpdateMessage"
        };
    }

    public render() {
        const {
            component
        } = this.state;

        return (
            <AdminLayout>
                <div style={{ display: "flex", flexFlow: "column", height: "100%" }}>
                    <div style={{ flex: "0 1 auto" }}>
                        <Text variant={'xxLargePlus'}>Mapping configuration</Text><br />
                        <Text variant={'small'}>
                            Configure mapping between Service Health Hub entity properties and task fields.
                        </Text><br />
                        <Text variant={'small'}>Select a service to edit mapping</Text><br />&nbsp;
                    </div>
                    <div style={{ flex: "1 1 auto" }}>
                        <Pivot
                            aria-label="Select service"
                            overflowBehavior="menu"
                            onLinkClick={this._handlePivotLinkClick}
                        >
                            <PivotItem headerText="Service Updates" itemKey="ServiceUpdateMessage">
                                {component === "ServiceUpdateMessage" ? (
                                    <ConfigEditor component='ServiceUpdateMessage' element='metadataMapping' />
                                ) : ("")}
                            </PivotItem>
                            <PivotItem headerText="Service Issues" itemKey="ServiceHealthIssue">
                                {component === "ServiceHealthIssue" ? (
                                    <ConfigEditor component='ServiceHealthIssue' element='metadataMapping' />
                                ) : ("")}
                            </PivotItem>
                            <PivotItem headerText="Microsoft 365 Roadmap" itemKey="RoadmapCommunication">
                                {component === "RoadmapCommunication" ? (
                                    <ConfigEditor component='RoadmapCommunication' element='metadataMapping' />
                                ) : ("")}
                            </PivotItem>
                            <PivotItem headerText="Service Health Hub Releases" itemKey="ReleaseMessage">
                                {component === "ReleaseMessage" ? (
                                    <ConfigEditor component='ReleaseMessage' element='metadataMapping' />
                                ) : ("")}
                            </PivotItem>
                            <PivotItem headerText="Azure Service Health" itemKey="AzureServiceHealthAlert">
                                {component === "AzureServiceHealthAlert" ? (
                                    <ConfigEditor component='AzureServiceHealthAlert' element='metadataMapping' />
                                ) : ("")}
                            </PivotItem>
                            <PivotItem headerText="Office 365 Endpoint change" itemKey="Office365EndpointsChange">
                                {component === "Office365EndpointsChange" ? (
                                    <ConfigEditor component='Office365EndpointsChange' element='metadataMapping' />
                                ) : ("")}
                            </PivotItem>
                        </Pivot>
                    </div>
                </div>
            </AdminLayout>
        );
    }

    private _handlePivotLinkClick = (item?: PivotItem) => {
        if (item) {
            this.setState({
                component: item.props.itemKey!
            });
        }
    };
}