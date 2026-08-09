export interface IServiceHealthEvent {
    id: string;
    title: string;
    classification: string;
    workload: string;
    internalWorkloadName: string;
    status: string;
    published: boolean;
}

export interface IServiceHealthEventPastAnalysis {
    advisories30: number;
    advisories60: number;
    incidents30: number;
    incidents60: number;
}

export interface IEventStatistics {
    advisories: string[];
    incidents: string[];
    impactedServices: string[];
    events: IServiceHealthEvent[];
    pastEvents?: IServiceHealthEventPastAnalysis;
}

export interface IRoadmapStatistics {
    inDevelopment: number;
    rollingOut: number;
    launched: number;
}

export interface IServiceStatistics {
    eventStatistics: IEventStatistics;
    roadmapStatistics: IRoadmapStatistics;
}

export interface IPublicIncident {
    id: string;
    title: string;
    service: string;
    startTime: Date;
    lastModified: Date;
    comments: string;
}

export interface IPublicMessage {
    id: string;
    title: string;
    services: string[];
    startTime: Date;
    lastModified: Date;
    comments: string;
    content: string;
}

export interface IPublicEvents {
    events: IPublicIncident[];
    messages: IPublicMessage[];
}