/*
    <copyright file="constants.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

export default class Constants {

	//Commands
	public static readonly SaveAdminDetailCommand: string = "SAVE ADMIN DETAILS";
	public static readonly UpdateAdminDetailCommand: string = "UPDATE ADMIN DETAILS";
	public static readonly SaveNominationCommand: string = "SAVE NOMINATED DETAILS";
	public static readonly NominateAwardsCommand: string = "NOMINATE AWARDS";
	public static readonly CancelCommand: string = "CANCEL";

	//Themes
	public static readonly body: string = "body";
	public static readonly theme: string = "theme";
	public static readonly default: string = "default";
	public static readonly light: string = "light";
	public static readonly dark: string = "dark";
	public static readonly contrast: string = "contrast";

	//KeyCodes
	public static readonly keyCodeEnter: number = 13;
    public static readonly keyCodeSpace: number = 32;

    public static readonly minimumCycleDays: number = 7;

}