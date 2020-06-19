/*
    <copyright file="nomination-award-preview.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class NominationAwardPreview {
    NominatedByName: string | "" = "";
    ImageUrl: string | "" = "";
    ReasonForNomination: string | undefined;
    Nominees: string[] = [];
    AwardId: string | undefined;
    AwardName?: string | undefined;
    TeamId: string | undefined;
    NomineeUserPrincipalNames: string[] = [];
    NomineeObjectIds: string[] = [];
    NominatedByUserPrincipalName: string | undefined;
    NominatedByObjectId?: string | null;
    telemetry?: any = null;
    locale?: string | null;
    theme?: string | null;
}