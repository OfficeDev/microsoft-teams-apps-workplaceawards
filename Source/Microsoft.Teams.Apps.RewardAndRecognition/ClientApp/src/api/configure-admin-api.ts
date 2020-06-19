/*
    <copyright file="configure-admin-api.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import axios from "./axios-decorator";
const baseAxiosUrl = window.location.origin;

/**
* Get all team members.
* @param  {String} teamId Team ID for getting members.
*/
export const getMembersInTeam = async (teamId: string | undefined): Promise<any> => {

    let url = baseAxiosUrl + `/api/ConfigureAdmin/teammembers?teamId=${teamId}`;
    return await axios.get(url);
}

/**
* Save admin details.
* @param  {AdminDetails | Null} adminDetails admin details.
*/
export const saveAdminDetails = async (adminDetails: any): Promise<any> => {

    let url = baseAxiosUrl + "/api/ConfigureAdmin/admindetail";
    return await axios.post(url, adminDetails);
}

/**
* Get team captain user detail.
* @param  {String} teamId Team ID for getting members.
*/
export const getUserRoleInTeam = async (teamId: string | undefined): Promise<any> => {

    let url = baseAxiosUrl + `/api/ConfigureAdmin/admindetail?teamId=${teamId}`;
    return await axios.get(url);
}

