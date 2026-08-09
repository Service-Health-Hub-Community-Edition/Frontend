import * as React from 'react';
import { Text, IStyleSet, Label, ILabelStyles, Pivot, PivotItem } from '@fluentui/react';
import { AdminLayout } from "./AdminLayout";
import { ConfigEditor } from "./jsonConfigEditor/ConfigEditor";

interface IAdminRoutingState {
    component: string
}

export class AdminRouting extends React.Component<{}, IAdminRoutingState> {
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
                <div style={{ display: "flex", flexFlow: "column", height: "100%"}}>
                    <div style={{ flex: "0 1 auto" }}>
                            <Text variant={'xxLargePlus'}>Routing configuration</Text><br />
                            <Text variant={'small'}>
                                Configure notification and task distribution by defining routing rules.<br />
                                Defining exclusion rules for unused services can help reduce amount of notifications and tasks.
                            </Text><br />
                            <Text variant={'small'}>Select a service to edit routing rules</Text><br />&nbsp;
                    </div>
                    <div style={{ flex: "1 1 auto" }}>
                        <Pivot
                            aria-label="Select service"
                            overflowBehavior="menu"
                            onLinkClick={this._handlePivotLinkClick}
                        >
                            <PivotItem headerText="Service Updates" itemKey="ServiceUpdateMessage">
                                {component === "ServiceUpdateMessage" ? (
                                    <ConfigEditor component='ServiceUpdateMessage' element='routing' />
                                ) : ("") }
                            </PivotItem>
                            <PivotItem headerText="Service Issues" itemKey="ServiceHealthIssue">
                                {component === "ServiceHealthIssue" ? (
                                    <ConfigEditor component='ServiceHealthIssue' element='routing' />
                                ) : ("")}
                            </PivotItem>
                            <PivotItem headerText="Microsoft 365 Roadmap" itemKey="RoadmapCommunication">
                                {component === "RoadmapCommunication" ? (
                                    <ConfigEditor component='RoadmapCommunication' element='routing' />
                                ) : ("")}
                            </PivotItem>
                            <PivotItem headerText="Service Health Hub Releases" itemKey="ReleaseMessage">
                                {component === "ReleaseMessage" ? (
                                    <ConfigEditor component='ReleaseMessage' element='routing' />
                                ) : ("")}
                            </PivotItem>
                            <PivotItem headerText="Azure Service Health" itemKey="AzureServiceHealthAlert">
                                {component === "AzureServiceHealthAlert" ? (
                                    <ConfigEditor component='AzureServiceHealthAlert' element='routing' />
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