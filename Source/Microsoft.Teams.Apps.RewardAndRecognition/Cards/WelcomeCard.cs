// <copyright file="WelcomeCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    ///  This class process welcome card when installed in team scope.
    /// </summary>
    public static class WelcomeCard
    {
        /// <summary>
        /// This method will construct the user welcome card when bot is added in team scope.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Welcome card.</returns>
        public static Attachment GetCard(string applicationBasePath, IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/AppIcon.png", applicationBasePath)),
                                        Size = AdaptiveImageSize.Large,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = localizer.GetString("WelcomeCardTitle"),
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Size = AdaptiveTextSize.Large,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = localizer.GetString("WelcomeTeamCardContent"),
                                        Wrap = true,
                                        Spacing = AdaptiveSpacing.None,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("WelcomeSubHeaderText"),
                        Spacing = AdaptiveSpacing.Small,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("SetAdminBulletPoint"),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("ManageAwardsBulletPoint"),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("SetRewardCycleBulletPoint"),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("NominateBulletPoint"),
                        Wrap = true,
                        Spacing = AdaptiveSpacing.None,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("ContentText"),
                        Spacing = AdaptiveSpacing.Small,
                    },
                },

                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SetAdminButtonText"),
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.FetchActionType,
                            },
                            Command = Constants.ConfigureAdminAction,
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
        }

        /// <summary>
        /// This method will construct the configure new admin card when existing award admin left the team.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Configure new admin card.</returns>
        public static Attachment ConfigureNewAdminCard(IStringLocalizer<Strings> localizer)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = localizer.GetString("WelcomeCardNewAdminText"),
                        Spacing = AdaptiveSpacing.Small,
                        Wrap = true,
                    },
                },

                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("SetAdminButtonText"),
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.FetchActionType,
                            },
                            Command = Constants.ConfigureAdminAction,
                        },
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
