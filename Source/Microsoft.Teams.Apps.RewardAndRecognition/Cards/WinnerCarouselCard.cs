// <copyright file="WinnerCarouselCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.RewardAndRecognition.Cards
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.CodeAnalysis;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Newtonsoft.Json;

    /// <summary>
    ///  This class process tour carousel feature to show winners.
    /// </summary>
    public static class WinnerCarouselCard
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
        /// Render the set of attachments that comprise carousel.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="winners">Award winner details.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>The card that comprise the winner details.</returns>
        public static IEnumerable<Attachment> GetAwardWinnerCard(string applicationBasePath, IEnumerable<AwardWinnerNotification> winners, IStringLocalizer<Strings> localizer)
        {
            var attachments = new List<Attachment>();
            foreach (var winner in winners.GroupBy(row => row.AwardId))
            {
                var groupNominations = winner.Select(rows => JsonConvert.DeserializeObject<List<string>>(rows.GroupName)).Distinct().ToList();
                string winnersName = string.Join(", ", groupNominations.SelectMany(row => row).ToList().Distinct().ToList());
                AdaptiveCard carouselCard = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
                {
                    Body = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Text = localizer.GetString("AwardWinnerCardTitle"),
                            Weight = AdaptiveTextWeight.Bolder,
                            Size = AdaptiveTextSize.Large,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = $"{localizer.GetString("WinnerCardRewardCycleTitle")}: {winner.First().AwardCycle}",
                            Size = AdaptiveTextSize.Small,
                            Spacing = AdaptiveSpacing.Small,
                            Wrap = true,
                        },
                        new AdaptiveImage
                        {
                            Url = string.IsNullOrEmpty(winner.First().AwardLink) ? new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/DefaultAwardImage.png", applicationBasePath)) : new Uri(winner.First().AwardLink),
                            PixelWidth = AwardImagePixelWidth,
                            PixelHeight = AwardImagePixelHeight,
                            Size = AdaptiveImageSize.Auto,
                            Style = AdaptiveImageStyle.Default,
                            HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = winner.OrderByDescending(row => row.NominatedOn).First().AwardName,
                            Size = AdaptiveTextSize.Large,
                            Weight = AdaptiveTextWeight.Bolder,
                            Spacing = AdaptiveSpacing.Small,
                            Wrap = true,
                        },
                        new AdaptiveTextBlock
                        {
                            Text = winnersName,
                            Size = AdaptiveTextSize.Small,
                            Spacing = AdaptiveSpacing.Medium,
                            Wrap = true,
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
