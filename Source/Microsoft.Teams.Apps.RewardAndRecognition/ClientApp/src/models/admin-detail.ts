/*
    <copyright file="admin-detail.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export class AdminDetails {
    CreatedByUserPrincipalName: string | "" = "";
    CreatedByObjectId?: string | null = null;
    CreatedOn: Date | null = null;
    AdminName: string | "" = "";
    AdminObjectId?: string | null = null;
    AdminUserPrincipalName: string | null = null;
    NoteForTeam: string | "" = "";
    TeamId: string | undefined;
}