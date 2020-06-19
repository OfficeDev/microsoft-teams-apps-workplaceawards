// <copyright file="awards-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Get all awards data from API
* @param {String} teamId Team Id for which the awards will be fetched
*/
export const getAllAwards = async (teamId: string): Promise<any> => {

    let url = baseAxiosUrl + `/api/Awards/allawards?teamId=${teamId}`;
    return await axios.get(url);
}

/**
* Get award data from API
* @param {String} teamId Team Id for which the awards will be fetched
*/
export const getAwardDetails = async (teamId: string | null, awardId: string | null): Promise<any> => {
    let url = baseAxiosUrl + `/api/Awards/awarddetails?teamId=${teamId}&awardId=${awardId}`;
    return await axios.get(url);
}

/**
* Post award data from API
* @param {String} teamId Team Id for which the awards will be fetched
*/
export const postAward = async (data: any): Promise<any> => {

    let url = baseAxiosUrl + "/api/Awards/award";
    return await axios.post(url, data);
}

/**
* Delete user selected award
* @param {string} awardIds selected award ids which needs to be deleted
* @param {string} teamId Team id
*/
export const deleteSelectedAwards = async (awardIds: string, teamId: string): Promise<any> => {

    let url = baseAxiosUrl + `/api/Awards/awards?teamId=${teamId}&awardIds=${awardIds}`;
    return await axios.delete(url);
}