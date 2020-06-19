// <copyright file="setting-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Get bot application settings from API.
*/
export const getBotSetting = async (teamId: string | undefined): Promise<any> => {

    let url = baseAxiosUrl + `/api/Settings/botsettings?teamId=${teamId}`;
    return await axios.get(url);
}