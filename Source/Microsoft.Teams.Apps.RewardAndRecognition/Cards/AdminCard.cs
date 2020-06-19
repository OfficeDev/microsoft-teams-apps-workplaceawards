// <copyright file="AdminCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    ///  This class process admin card when configured.
    /// </summary>
    public static class AdminCard
    {
        /// <summary>
        /// This method will construct admin card with corresponding details.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="adminDetails">Admin details to show in card.</param>
        /// <returns>User welcome card.</returns>
        public static Attachment GetAdminCard(IStringLocalizer<Strings> localizer, TaskModuleResponseDetails adminDetails)
        {
            adminDetails = adminDetails ?? throw new ArgumentNullException(nameof(adminDetails));

            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("AdminHeaderText"),
                        Weight = AdaptiveTextWeight.Bolder,
                        Size = AdaptiveTextSize.Large,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("AdminSubheaderText", adminDetails.AdminName, adminDetails.AdminUserPrincipalName),
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("NoteForTeamText", adminDetails.NoteForTeam),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.Default,
                        IsVisible = !string.IsNullOrEmpty(adminDetails.NoteForTeam),
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("AdminTabNavigationText", adminDetails.NoteForTeam),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.Large,
                        Size = AdaptiveTextSize.Medium,
                        Weight = AdaptiveTextWeight.Bolder,
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }
    }
}
