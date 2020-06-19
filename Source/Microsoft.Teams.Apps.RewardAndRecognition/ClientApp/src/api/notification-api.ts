// <copyright file="notification-api.ts" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import axios from "./axios-decorator";

const baseAxiosUrl = window.location.origin;

/**
* Send reward winner notification from API.
* @param {object} data winner awards card details.
*/
export const sendWinnerNotification = async (data: any): Promise<any> => {

    let url = baseAxiosUrl + "/api/Notification/winnernotification";
    return await axios.post(url, data);
}