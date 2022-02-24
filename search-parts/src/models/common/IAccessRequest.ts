export interface IAccessRequest {
    Title?: string;
    ReportNumber?: string;
    RequestAccessStatus?: RequestAccessStatus,
    Author?: string;
    Publisher?: string;
    ReportComments?: string;
    UserComments?: string;
    ReportPath?: string;    
    ReportClassification?: ReportClassification;
}

export interface IAccessRequestResults {
    ID?: number;
    Id?: number;
    Title?: string;
    ReportNumber?: string;
    RequestAccessStatus?: RequestAccessStatus,
    Requester?: {
        ID?: number;
        Title?: string
    }
}

export enum RequestAccessStatus {
    New,
    InProgress,
    Completed
}

export enum ReportClassification {
    Confidential,
    Restricted,
    Unrestricted
}

export enum IAccessRequestResultsType {
    Success,
    Failure
}