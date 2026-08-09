import * as React from 'react';
import * as msTeams from '@microsoft/teams-js';
import { Text } from '@fluentui/react';

export class Configure extends React.Component<{}> {
    constructor(props: {}) {
        super(props);

    }

    public render() {
        msTeams.settings.registerOnSaveHandler(saveEvent => {
            msTeams.settings.setSettings({
                contentUrl: window.location.origin,
                entityId: window.location.origin
            });

            saveEvent.notifySuccess();
        });
        msTeams.settings.setValidityState(true);

        return (
            <div>
                <Text variant={'xxLargePlus'}>Welcome to the Service Health Dashboard App</Text><br />
                <Text variant={'small'}>Click <b>Save</b> to save the configuration and add the app to the channel.</Text><br />&nbsp;
            </div>
        );
    }
}