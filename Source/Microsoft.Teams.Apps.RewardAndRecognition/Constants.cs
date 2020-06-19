// <copyright file="Constants.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition
{
    /// <summary>
    /// Constant values that are used in multiple files.
    /// </summary>
    public static class Constants
    {
        /// <summary>
        /// Describes adaptive card version to be used. Version can be upgraded or changed using this value.
        /// </summary>
        public const string AdaptiveCardVersion = "1.2";

        /// <summary>
        /// Date time format to support adaptive card text feature.
        /// </summary>
        /// <remarks>
        /// refer adaptive card text feature https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/text-features#datetime-formatting-and-localization.
        /// </remarks>
        public const string Rfc3339DateTimeFormat = "yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'";

        /// <summary>
        /// Message back card action.
        /// </summary>
        public const string MessageBackActionType = "messageBack";

        /// <summary>
        /// Task fetch action type.
        /// </summary>
        public const string FetchActionType = "task/fetch";

        /// <summary>
        /// Set champion action.
        /// </summary>
        public const string ConfigureAdminAction = "CONFIGURE ADMIN";

        /// <summary>
        /// Save admin details action.
        /// </summary>
        public const string SaveAdminDetailsAction = "SAVE ADMIN DETAILS";

        /// <summary>
        /// Update admin details action.
        /// </summary>
        public const string UpdateAdminDetailCommand = "UPDATE ADMIN DETAILS";

        /// <summary>
        /// Save nomination detail action.
        /// </summary>
        public const string SaveNominatedDetailsAction = "SAVE NOMINATED DETAILS";

        /// <summary>
        /// Manage award action.
        /// </summary>
        public const string ManageAwardAction = "MANAGE AWARDS";

        /// <summary>
        /// Nominate awards action.
        /// </summary>
        public const string NominateAction = "NOMINATE AWARDS";

        /// <summary>
        /// Endorse award action.
        /// </summary>
        public const string EndorseAction = "ENDORSE AWARD";

        /// <summary>
        /// Cancel command.
        /// </summary>
        public const string CancelCommand = "CANCEL";

        /// <summary>
        /// Ok command.
        /// </summary>
        public const string OkCommand = "OK";

        /// <summary>
        /// Nominate award table.
        /// </summary>
        public const string NominateAwardTable = "NominationDetail";

        /// <summary>
        /// default value for channel activity to send notifications.
        /// </summary>
        public const string TeamsBotFrameworkChannelId = "msteams";
    }
}
