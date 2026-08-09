import * as React from 'react';
import { Text, Separator, Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';
import { ServiceComponent } from '../../../ServiceNameComponent';
import { acquireAccessToken } from "../../../../auth/AccessTokenHelper";
import { GlobalState } from "../../../GlobalState";
import { AuthenticationResult } from '@azure/msal-browser';
import { translate, getUIString } from "../../../../translation/translator";

export interface IMessageDetailsState {
    id: string;
    title: string;
    service: string[];
    latestMessage: string;
    comments: string;
    lastUpdated: Date;
    uiStrings: Map<string, string>;
    error?: string;
}

export class PublicMessageDetails extends React.Component<{ id?: string }, IMessageDetailsState> {
    static contextType = GlobalState;

    constructor(props: { id?: string }) {
        super(props);

        var uiStrings: Map<string, string> = new Map([
            ["Last updated", "Last updated"],
            ["Service", "Service"],
            ["IT department comments", "IT department comments"],
            ["Message summary", "Message summary"]
        ]);

        this.state = {
            id: "",
            title: "",
            service: [],
            latestMessage: "",
            comments: "",
            lastUpdated: new Date(),
            uiStrings: uiStrings,
            error: undefined
        };
    }

    public render() {
        const {
            id, service, latestMessage, comments, lastUpdated, uiStrings, error } = this.state;

        return (
            <div>
                <div style={{ display: error != undefined ? 'block' : 'none' }}>
                    <MessageBar
                        messageBarType={MessageBarType.error}
                        isMultiline={false}
                    >
                        Couldn't retrieve data. Error: {error}
                    </MessageBar>
                    <br />
                </div>
                <div className="messageDetails" style={{ display: id && error == undefined ? 'block' : 'none' }}>
                    <Separator />
                    <Text>
                        <b>{id}</b> • <b>{getUIString(uiStrings, "Last updated") }:</b> {lastUpdated.toLocaleString()}
                    </Text>
                    <Separator />
                    <Text variant={'medium'} block>
                        <b>{getUIString(uiStrings, "Service")}</b>
                    </Text>
                    <Text>
                        {service.map((s) => (
                            <ServiceComponent name={s} />
                            ))}
                    </Text>

                    {comments.trim().length > 0 ? (
                        <div>
                            <Text variant={'medium'} block>
                                <br /><b>{getUIString(uiStrings, "IT department comments")}</b>
                            </Text>
                            <Text>
                                {comments}
                            </Text>
                        </div>
                    ) : ""}


                    <Text variant={'medium'} block>
                        <Separator />
                        <b>{getUIString(uiStrings, "Message summary")}</b><br/>
                    </Text>
                    <Text>
                        <div dangerouslySetInnerHTML={{ __html: latestMessage }}/>
                    </Text>
                    <Separator />
                </div>

                <div className="loadingProgress" style={{ display: id ? 'none' : 'block' }}>
                    <br/><br/>
                    <Spinner size={SpinnerSize.large} />
                </div>
            </div>
        );
    }

    componentDidMount() {
        this._getMessageDetails(this.props.id);
    }

    componentDidUpdate(prevProps: { id?: string }) {
        if (prevProps.id !== this.props.id) {
            this._getMessageDetails(this.props.id);
        }
    }

    private _getMessageDetails(id?: string) {
        var messageId = id;
        let globalState: any = this.context;
        var language: string = globalState.getProfileProperty("language", "en").trim();
        var translatedUIStrings: Map<string, string>;
        var langQuery: string = "";
        if (language !== "" && language !== "en") {
            langQuery = "&lang=" + language;
        }

        var authResponse: AuthenticationResult;

        if (messageId != undefined && messageId.trim().length > 0) {
            acquireAccessToken()
                .then((response) => {
                    authResponse = response;

                    translate(language,
                        this.state.uiStrings,
                        authResponse
                    ).then((response) => {                      
                        translatedUIStrings = response;
                    });
                }).then(() => {
                    fetch('/api/PublicMessage?id=' + messageId + langQuery, { headers: authResponse.idToken === "" ? {} : { 'Authorization': `Bearer ${authResponse.idToken}` } })
                        .then(response => {
                            if (response.ok) {
                                return response.json();
                            } else {
                                this.setState({
                                    error: response.status + " " + response.statusText
                                });
                                throw Error(response.status + " " + response.statusText);
                            }
                        })
                        .then(r => {
                            var result = r;
                            if (result != null) {
                                this.setState({
                                    id: result.id,
                                    title: result.title,
                                    service: result.services.length > 0 ? result.services : ["General announcement"],
                                    latestMessage: result.content.replaceAll(/\[(.*?)\]/g, '<b>$1</b>'),
                                    lastUpdated: new Date(result.lastModified),
                                    uiStrings: translatedUIStrings,
                                    comments: result.comments
                                });
                            }
                        });
                }).catch((err) => {
                    this.setState({
                        error: err.message
                    });
                });
        }
    }
}