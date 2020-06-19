import { getMembersInTeam } from "../api/configure-admin-api";

export const isNullorWhiteSpace = (input: string): boolean => {
    return !input || !input.trim();
}

export const checkUrl = (url: string) => {
    return (url.match(/\.(jpeg|jpg|gif|png)$/) != null);
}

export const getBaseUrl = () => {
    return window.location.origin;
}

/**
    *Navigate to error page
*/
export const navigateToErrorPage = async (code: string) => {
    return window.location.href = `/errorpage?code=${code}`;
}

/**
    *validate user is part of team.
*/
export const validateUserPartOfTeam = async (teamId: string, userObjectId: string) => {
    let teamMembers = await getMembersInTeam(teamId);
    if (teamMembers.status === 200 && teamMembers.data) {
        let member = teamMembers.data.find(element => element.aadobjectid === userObjectId);
        if (member !== null || member !== undefined) {
            return true;
        }
        else {
            return false;
        }
    }
    else {
        navigateToErrorPage(teamMembers.status);
    }
}