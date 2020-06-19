// <copyright file="NominateCarouselCard.cs" company="Microsoft">
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
    ///  This class process award nomination carousel cards.
    /// </summary>
    public static class NominateCarouselCard
    {
        /// <summary>
        /// Represents the height of award image in pixel.
        /// </summary>
        private const int AwardImagePixelHeight = 243;

        /// <summary>
        /// Represents the width of award image in pixel.
        /// </summary>
        private const int AwardImagePixelWidth = 243;

        /// <summary>
        /// Get award nomination carousel cards.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="awards">award details.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="details">Details to show in card.</param>
        /// <returns>The cards that comprise nominations.</returns>
        public static IEnumerable<Attachment> GetAwardNominationCards(string applicationBasePath, IEnumerable<AwardEntity> awards, IStringLocalizer<Strings> localizer, TaskModuleResponseDetails details)
        {
            awards = awards ?? throw new ArgumentNullException(nameof(awards));

            var attachments = new List<Attachment>();
            var startCycleDate = "{{DATE(" + details?.RewardCycleStartDate.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture) + ", SHORT)}}";
            var endCycleDate = "{{DATE(" + details?.RewardCycleEndDate.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture) + ", SHORT)}}";

            foreach (var award in awards)
            {
                AdaptiveCard carouselCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
                {
                    Body = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = localizer.GetString("RewardTitle"),
                            Weight = AdaptiveTextWeight.Bolder,
                            Size = AdaptiveTextSize.Large,
                        },
                        new AdaptiveImage
                        {
                            Url = string.IsNullOrEmpty(award.AwardLink) ? new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/DefaultAwardImage.png", applicationBasePath)) : new Uri(award.AwardLink),
                            PixelWidth = AwardImagePixelWidth,
                            PixelHeight = AwardImagePixelHeight,
                            Size = AdaptiveImageSize.Auto,
                            Style = AdaptiveImageStyle.Default,
                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = $"**{award.AwardName.Trim()}**",
                            Size = AdaptiveTextSize.Large,
                            Weight = AdaptiveTextWeight.Bolder,
                            Spacing = AdaptiveSpacing.Small,
                            Wrap = true,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = localizer.GetString("RewardCycleHeader", startCycleDate, endCycleDate),
                            Size = AdaptiveTextSize.Small,
                            Spacing = AdaptiveSpacing.Small,
                            Wrap = true,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = award.AwardDescription,
                            Size = AdaptiveTextSize.Small,
                            Spacing = AdaptiveSpacing.Small,
                            Wrap = true,
                        },
                    },
                    Actions = new List<AdaptiveAction>
                    {
                        new AdaptiveSubmitAction
                        {
                            Title = localizer.GetString("NominateButtonText"),
                            Data = new AdaptiveCardAction
                            {
                                MsteamsCardAction = new CardAction
                                {
                                    Type = Constants.FetchActionType,
                                },
                                Command = Constants.NominateAction,
                                AwardId = award.AwardId,
                                RewardCycleId = details.RewardCycleId,
                            },
                        },
                    },
                };

                attachments.Add(new Attachment
                {
                    ContentType = AdaptiveCard.ContentType,
                    Content = carouselCard,
                });
            }

            return attachments;
        }
    }
}
