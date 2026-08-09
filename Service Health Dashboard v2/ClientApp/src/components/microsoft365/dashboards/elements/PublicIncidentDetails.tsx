import * as React from 'react';
import { Text, Separator } from '@fluentui/react';
import { IPublicIncident } from '../../../IServiceStatistics';
import { GlobalState } from '../../../GlobalState';
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";
import { translate, getUIString } from "../../../../translation/translator";

export interface IPublicIncidentDetailsState {
    Id: string;
    Title: string;
    Service: string;
    LastModified: Date;
    StartTime: Date;
    Comments: string;
    uiStrings: Map<string, string>;
}

export class PublicIncidentDetails extends React.Component<{ incidentDetails?: IPublicIncident }, IPublicIncidentDetailsState> {
    static contextType = GlobalState;

    constructor(props: { incidentDetails?: IPublicIncident }) {
        super(props);

        var uiStrings: Map<string, string> = new Map([
            ["Last updated", "Last updated"],
            ["Start time", "Start time"],
            ["Comment", "Comment"],
            ["Message summary", "Message summary"]
        ]);

        if (this.props.incidentDetails === undefined) {
            this.state = {
                Id: "",
                Title: "",
                Service: "",
                LastModified: new Date(),
                StartTime: new Date(),
                Comments: "",
                uiStrings: uiStrings
            };
        } else {
            this.state = {
                Id: this.props.incidentDetails.id,
                Title: this.props.incidentDetails.title,
                Service: this.props.incidentDetails.service,
                LastModified: new Date(this.props.incidentDetails.lastModified),
                StartTime: new Date(this.props.incidentDetails.startTime),
                Comments: this.props.incidentDetails.comments,
                uiStrings: uiStrings
            };
        }
    }

    public render() {
        const {
            Id, Service, Comments, StartTime, LastModified, uiStrings } = this.state;

        return (
            <div>
                <div className="PublicIncidentDetails">
                    <Separator />
                    <Text>
                        <b>{Id}</b> • {Service} • <b>{getUIString(uiStrings, "Start time")}:</b> {StartTime.toLocaleString()} • <b>{getUIString(uiStrings, "Last updated")}:</b> {LastModified.toLocaleString()}
                    </Text>
                    <Separator />
                    <Text variant={'medium'} block>
                        <b>{getUIString(uiStrings, "Comment")}</b>
                    </Text>
                    <Text>
                        <div style={{ whiteSpace: "pre-wrap" }} dangerouslySetInnerHTML={{ __html: Comments }}/>
                    </Text>
                </div>
            </div >
        );
    }

    componentDidMount() {
        let globalState: any = this.context;
        var language: string = globalState.getProfileProperty("language", "en").trim();
        var translatedUIStrings: Map<string, string>;    

        acquireAccessToken()
                .then((response: any) => {
                    translate(language,
                        this.state.uiStrings,
                        response
                    ).then((response: any) => {                      
                        translatedUIStrings = response;
                        this.setState({
                            uiStrings: translatedUIStrings
                        });
                    });
                })  
    }

    componentDidUpdate(prevProps: { incidentDetails: IPublicIncident }) {
        if (prevProps.incidentDetails?.id !== this.props.incidentDetails?.id) {
            if (this.props.incidentDetails === undefined) {
                this.setState({
                    Id: "",
                    Title: "",
                    Service: "",
                    LastModified: new Date(),
                    StartTime: new Date(),
                    Comments: "",
                });
            } else {
                this.setState({
                    Id: this.props.incidentDetails.id,
                    Title: this.props.incidentDetails.title,
                    Service: this.props.incidentDetails.service,
                    LastModified: new Date(this.props.incidentDetails.lastModified),
                    StartTime: new Date(this.props.incidentDetails.startTime),
                    Comments: this.props.incidentDetails.comments
                });
            }
        }
    }
}