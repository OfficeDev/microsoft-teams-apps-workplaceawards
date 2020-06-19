/*
    <copyright file="nominate-entity.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class NominateEntity {
    AwardId: string | undefined;
    RewardCycleId: string | undefined;
    AwardName?: string | undefined;
    ReasonForNomination?: string | undefined;
    TeamId: string | undefined;
    NominatedOn: Date | undefined;
    NomineeNames?: string | undefined;
    NomineeUserPrincipalNames: string | undefined;
    NomineeObjectIds: string | undefined;
    NominatedByName?: string | undefined;
    NominatedByUserPrincipalName: string | undefined;
    NominatedByObjectId?: string | null;
    GroupName?: string | undefined;
    AwardImageLink?: string | undefined;
}