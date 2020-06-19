// <copyright file="reward-cycle-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Set reward cycle from API.
* @param {String} teamId Team Id for which the awards will be fetched.
*/
export const setRewardCycle = async (data: any): Promise<any> => {

    let url = baseAxiosUrl + "/api/RewardCycle/rewardcycle";
    return await axios.post(url, data);
}

/**
* Get reward cycle from API.
* @param {String} teamId Team Id for which the awards will be fetched.
* @param {boolean} isActiveCycle Flag to identify active cycle or published cycle.
*/
export const getRewardCycle = async (teamId: string | undefined, isActiveCycle: boolean | true): Promise<any> => {

    let url = baseAxiosUrl + `/api/RewardCycle/rewardcycledetails?teamId=${teamId}&isActiveCycle=${isActiveCycle}`;
    return await axios.get(url);
}