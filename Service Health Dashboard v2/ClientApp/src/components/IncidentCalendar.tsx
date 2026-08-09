import * as React from 'react';
import { ConstrainMode, Text } from '@fluentui/react';
import { DetailsList, ShimmeredDetailsList, DetailsListLayoutMode, Selection, SelectionMode, IColumn } from '@fluentui/react';
import { DefaultButton, PrimaryButton, Dialog, DialogType, DialogFooter } from '@fluentui/react';
import { Link, ILinkStyleProps, ILinkStyles } from '@fluentui/react';
import { Modal, Spinner, SpinnerSize } from '@fluentui/react';
import RevoCalendar from 'revo-calendar';
import 'revo-calendar/dist/index.css';
import { IncidentDetails } from './IncidentDetails';
import { MessageBar, MessageBarType } from '@fluentui/react';
import { acquireAccessToken } from "../auth/AccessTokenHelper";

interface IEventExtraInfo {
    icon: string;
    text: string;
}

interface IEvents {
    name: string;
    date: number;
    allday: boolean;
    extra: IEventExtraInfo;
}

interface IEventDetails {
    id: string;
    title: string;
    service: string;
    start: Date;
    end?: Date;
    duration: string;
}

interface IDate {
    day: number;
    month: number;
    year: number;
}

export interface IIncidentCalendarState {
    filter: string[];
    items: IEvents[];
    detailItems: IEventDetails[];
    columns: IColumn[];
    selectedDate: Date;
    initialized: boolean;
    isMessageDetailsDialogOpen: boolean;
    selectedMessage: string;
    error?: string;
}

var _initialized: boolean = false;

export class IncidentCalendar extends React.Component<{}, IIncidentCalendarState> {
    private _allItems: IEvents[];
    private _allDetails: IEventDetails[];

    constructor(props:{}) {
        super(props);

        this._allItems = [];
        this._allDetails = [];

        const pipeFabricStyles = (p: ILinkStyleProps): ILinkStyles => ({
            root: {
                textDecoration: 'none',
                color: p.theme.semanticColors.bodyText,
                fontWeight: '600',
                fontSize: p.theme.fonts.medium.fontSize,
            },
        });

        const columns: IColumn[] = [
            {
                key: 'clId',
                name: 'Id',
                fieldName: 'id',
                minWidth: 80,
                maxWidth: 80,
                isRowHeader: true,
                isResizable: false,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clTitle',
                name: 'Title',
                fieldName: 'title',
                minWidth: 180,
                maxWidth: 200,
                isResizable: false,
                isCollapsible: false,
                isMultiline: true,
                data: 'string',
                onRender: (item: IEventDetails) => {
                    return <Link onClick={(event) => {
                        event.preventDefault();
                        this._onOpenMessageDetails(item.id);
                    }} styles={pipeFabricStyles}>{item.title}</Link >;
                },
                isPadded: true,
            },

            {
                key: 'clService',
                name: 'Service',
                fieldName: 'service',
                minWidth: 60,
                maxWidth: 80,
                isResizable: true,
                isCollapsible: true,
                isMultiline: true,
                data: 'string',
                isPadded: true,
            },
            {
                key: 'clDuration',
                name: 'Duration',
                fieldName: 'duration',
                minWidth: 80,
                maxWidth: 100,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                isPadded: true,
            }
        ];

        this.state = {
            filter: [],
            items: [],
            detailItems: [],
            columns: columns,
            selectedDate: new Date(),
            initialized: false,
            isMessageDetailsDialogOpen: false,
            selectedMessage: "",
            error: undefined
        };
    }

    public render() {
        const { items, columns, detailItems, isMessageDetailsDialogOpen, selectedMessage, initialized, error } = this.state;

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

                <RevoCalendar
                    // primaryColor="#0078D4"
                    primaryColor="#323130"
                    secondaryColor="#ffffff"
                    textColor="#323130"
                    todayColor="#f3f2f1"
                    indicatorColor="red"
                    allowAddEvent={false}
                    allowDeleteEvent={false}
                    sidebarDefault={false}
                    detailDefault={false}
                    showDetailToggler={false}
                    openDetailsOnDateSelection={false}
                    events={items}
                    dateSelected={(date: IDate) => this._dateSelected(date)}
                />

                <DetailsList
                    items={detailItems}
                    compact={false}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                    isHeaderVisible={true}
                />

                <Dialog
                    hidden={!isMessageDetailsDialogOpen}
                    onDismiss={this._onDismisMessageDetails}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Message details',
                    }}
                    modalProps={{
                        /* titleAriaId: this._labelId,
                        subtitleAriaId: this._subTextId, */
                        isBlocking: false
                    }}
                    styles={{
                        main: {
                            selectors: {
                                ['@media (min-width: 480px)']: {
                                    width: 720,
                                    minWidth: 300,
                                    maxWidth: '1000px'
                                }
                            }
                        }
                    }}
                >
                    <IncidentDetails id={selectedMessage} />

                    <DialogFooter>
                        <PrimaryButton onClick={this._onDismisMessageDetails} text="Close" />
                    </DialogFooter>
                </Dialog>

                <Modal
                    titleAriaId="Loading"
                    isOpen={!initialized}
                    isBlocking={false}
                >
                    <br /><br /><br />
                    <Spinner size={SpinnerSize.large} />
                    <br /><br /><br />

                </Modal>
            </div>

        );
    }

    componentDidMount() {
        this._getIncidents();
    }

    componentDidUpdate() {
        this._getIncidents();
    }

    private _getIncidents() {
        var items: IEvents[] = [];
        var details: IEventDetails[] = [];

        acquireAccessToken()
            .then((response) => {
                fetch('/api/Issues?select=id,title,startDateTime,endDateTime,service',
                    { headers: response.idToken === "" ? {} : { 'Authorization': `Bearer ${response.idToken}` } })
                    .then(response => {
                        if (response.ok) {
                            return response.json();
                        } else {
                            var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                            this.setState({
                                error: errorMsg + "Fetching data from /api/Issues endpoint failed with an error " + response.status + " " + response.statusText
                            });
                            throw Error(response.status + " " + response.statusText);
                        }
                    })
                .then(result => {
                    for (const incident of result) {
                        items.push({
                            name: incident.id + " " + incident.service,
                            date: +(new Date(incident.startDateTime)),
                            allday: true,
                            extra: {
                                icon: "",
                                text: incident.title
                            }
                        });

                        var start: Date = new Date(incident.startDateTime);
                        var end: Date | undefined = incident.endDateTime ? new Date(incident.endDateTime) : undefined;
                        var durationNum = end ? end.valueOf() - start.valueOf() : -1;
                        var durationDays = Math.trunc(durationNum / (1000 * 60 * 60 * 24));
                        var duration: string = end ? (durationDays > 0 ? durationDays + "d " : "") + new Intl.DateTimeFormat(
                            window.navigator.language, { hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).format(durationNum) : "ongoing";

                        details.push({
                            id: incident.id,
                            title: incident.title,
                            service: incident.service,
                            start: start,
                            end: end,
                            duration: duration
                        });
                    }
                }).then(() => {
                    this._allItems = items;
                    this._allDetails = details;
                    this.setState({
                        items: this._allItems,
                        detailItems: this._filterDetails(),
                        initialized: true
                    });
                    _initialized = true;
                });
            }).catch((err) => {
                var errorMsg: string = this.state.error === undefined ? "" : this.state.error + "<br/>";
                this.setState({
                    error: errorMsg + "Failed retrieving data from /api/Issues with an error " + err.message
                });
            });
    }

    private _filterDetails() {
        var selectedDate: Date = this.state.selectedDate;

        return this._allDetails.filter(event =>
            event.start.getDate() == selectedDate.getDate() &&
            event.start.getMonth() == selectedDate.getMonth() &&
            event.start.getFullYear() == selectedDate.getFullYear());
    }

    private _dateSelected(date: IDate) {
        if (_initialized) {
            var sDate: Date = new Date(date.year, date.month, date.day);

            this.setState({
                selectedDate: sDate
            });

            this.setState({
                detailItems: this._filterDetails()
            });
        }
    }

    private _onOpenMessageDetails(id: string): void {
        this.setState({
            selectedMessage: id,
            isMessageDetailsDialogOpen: true
        });
    }

    private _onDismisMessageDetails = (): void => {
        this.setState({
            selectedMessage: "",
            isMessageDetailsDialogOpen: false
        });
    }
}