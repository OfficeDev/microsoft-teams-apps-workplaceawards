/*
    <copyright file="result.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class ResultDetails {
    AwardId: string | undefined;
    AwardName: string | undefined;
    WinnerCount: number | undefined;
    NominationId: string | undefined;
    TeamId: string | undefined;
    NomineeNames: string | undefined;
    NomineeObjectIds: string | undefined;
    NomineeUserPrincipalNames: string | undefined;
    AwardLink: string | undefined;
    AwardCycle: string | undefined;
    GroupName: Array<string> | undefined;
}

export class NominatedAward {
    AwardId: string | undefined;
    AwardName: string | undefined;
}