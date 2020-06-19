/*
    <copyright file="nominate-awards-api.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import axios from "./axios-decorator";
const baseAxiosUrl = window.location.origin;

/**
* Save nominated details.
* @param  {NominateEntity | Null} nominateDetails nominated details.
*/
export const saveNominateDetails = async (nominateDetails: any): Promise<any> => {

    let url = baseAxiosUrl + "/api/Nominations/nomination";
    return await axios.post(url, nominateDetails);
}

/**
* Check duplicate award nomination.
* @param  {String | Null} teamId Team id.
* @param  {String | Null} aadObjectId User azure active directory object id.
* @param  {String | Null} cycleId Active award cycle unique id.
*/
export const checkDuplicateNomination = async (teamId: string | null, aadObjectIds: string | null, cycleId: string | null, awardId: string, nominatedByObjectId: string): Promise<any> => {
    let url = baseAxiosUrl + `/api/Nominations/checkduplicatenomination?teamId=${teamId}&aadObjectIds=${aadObjectIds}&cycleId=${cycleId}&awardId=${awardId}&nominatedByObjectId=${nominatedByObjectId}`;
    return await axios.get(url);
}

/**
* Get all nominations from API.
* @param {String} teamId Team Id for which the awards will be fetched.
 *@param {boolean} isAwardGranted flag: true for published award, else false.
 *@param {String} awardCycleId Active award cycle unique id.
*/
export const getAllAwardNominations = async (teamId: string | undefined, isAwardGranted: boolean | undefined, awardCycleId: string | undefined): Promise<any> => {

    let url = baseAxiosUrl + `/api/Nominations/allnominations?teamId=${teamId}&isAwardGranted=${isAwardGranted}&awardCycleId=${awardCycleId}`;
    return await axios.get(url);
}

/**
* publish nominations from API
* @param {String} teamId Team Id for which the awards will be fetched.
 *@param {String} nominationIds Publish nomination ids.
*/
export const publishAwardNominations = async (teamId: string | undefined, nominationIds: string | undefined): Promise<any> => {

    let url = baseAxiosUrl + `/api/Nominations/publishnominations?teamId=${teamId}&nominationIds=${nominationIds}`;
    return await axios.get(url);
}