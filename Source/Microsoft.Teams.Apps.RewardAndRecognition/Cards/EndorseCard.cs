// <copyright file="EndorseCard.cs" company="Microsoft">
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
    using Newtonsoft.Json;

    /// <summary>
    ///  This class process endorse card when award is nominated.
    /// </summary>
    public static class EndorseCard
    {
        /// <summary>
        /// Represents the height of default award image in pixel.
        /// </summary>
        private const int AwardImagePixelHeight = 85;

        /// <summary>
        /// Represents the width of default award image in pixel.
        /// </summary>
        private const int AwardImagePixelWidth = 85;

        /// <summary>
        /// Represents the height of endorse icon in pixel.
        /// </summary>
        private const int SuccessIconPixelHeight = 15;

        /// <summary>
        /// Represents the width of endorse icon in pixel.
        /// </summary>
        private const int SuccessIconPixelWidth = 15;

        /// <summary>
        /// Represents the container minimum height in pixel.
        /// </summary>
        private const int ContainerPixelHeight = 144;

        /// <summary>
        /// This method will construct endorse card with corresponding details.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="nominatedDetails">Nominated details to show in card.</param>
        /// <returns>Endorse card with nominated details.</returns>
        public static Attachment GetEndorseCard(string applicationBasePath, IStringLocalizer<Strings> localizer, TaskModuleResponseDetails nominatedDetails)
        {
            nominatedDetails = nominatedDetails ?? throw new ArgumentNullException(nameof(nominatedDetails));

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
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = nominatedDetails.AwardName,
                                        Wrap = true,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Size = AdaptiveTextSize.Large,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = string.IsNullOrEmpty(nominatedDetails.AwardLink) ? new Uri(string.Format(CultureInfo.InvariantCulture, "{0}/Content/DefaultAwardImage.png", applicationBasePath)) : new Uri(nominatedDetails.AwardLink),
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        PixelHeight = AwardImagePixelHeight,
                                        PixelWidth = AwardImagePixelWidth,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        Text = string.Join(", ", JsonConvert.DeserializeObject<List<string>>(nominatedDetails.GroupName)),
                        Wrap = true,
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Weight = AdaptiveTextWeight.Bolder,
                        Spacing = AdaptiveSpacing.Large,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = localizer.GetString("NominatedByText", nominatedDetails.NominatedByName),
                        Wrap = true,
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Spacing = AdaptiveSpacing.Default,
                    },
                    new AdaptiveTextBlock
                    {
                        Text = nominatedDetails.ReasonForNomination,
                        Wrap = true,
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Spacing = AdaptiveSpacing.Default,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("EndorseButtonText"),
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.FetchActionType,
                            },
                            Command = Constants.EndorseAction,
                            NomineeUserPrincipalNames = nominatedDetails.NomineeUserPrincipalNames,
                            AwardName = nominatedDetails.AwardName,
                            NomineeNames = nominatedDetails.NomineeNames,
                            NomineeObjectIds = nominatedDetails.NomineeObjectIds,
                            AwardId = nominatedDetails.AwardId,
                            RewardCycleId = nominatedDetails.RewardCycleId,
                            GroupName = string.Join(", ", JsonConvert.DeserializeObject<List<string>>(nominatedDetails.GroupName)),
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
        /// Construct the card to render endorse message to task module.
        /// </summary>
        /// <param name="applicationBasePath">Application base URL.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="awardName">Award name.</param>
        /// <param name="nomineeNames">Nominated users.</param>
        /// <param name="rewardCycleEndDate">Cycle end date.</param>
        /// <param name="isEndorsementSuccess">Gets the endorsement status.</param>
        /// <returns>Card attachment.</returns>
        public static Attachment GetEndorseStatusCard(string applicationBasePath, IStringLocalizer<Strings> localizer, string awardName, string nomineeNames, DateTime rewardCycleEndDate, bool isEndorsementSuccess)
        {
            var endCycleDate = "{{DATE(" + rewardCycleEndDate.ToString(Constants.Rfc3339DateTimeFormat, CultureInfo.InvariantCulture) + ", SHORT)}}";
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveContainer()
                    {
                        PixelMinHeight = ContainerPixelHeight,
                        Items = new List<AdaptiveElement>()
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
                                                Spacing = AdaptiveSpacing.Large,
                                                PixelWidth = SuccessIconPixelWidth,
                                                PixelHeight = SuccessIconPixelHeight,
                                                Url = new Uri(string.Format(CultureInfo.InvariantCulture, isEndorsementSuccess ? "{0}/Content/SuccessIcon.png" : "{0}/Content/ErrorIcon.png", applicationBasePath)),
                                            },
                                        },
                                    },
                                    new AdaptiveColumn
                                    {
                                        Width = AdaptiveColumnWidth.Stretch,
                                        Height = AdaptiveHeight.Auto,
                                        Items = new List<AdaptiveElement>
                                        {
                                            new AdaptiveTextBlock
                                            {
                                                Text = isEndorsementSuccess ? localizer.GetString("SuccessfulEndorseMessage", nomineeNames, awardName, endCycleDate) : localizer.GetString("AlreadyEndorsedMessage", nomineeNames),
                                                Wrap = true,
                                                Size = AdaptiveTextSize.Default,
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = localizer.GetString("OkButtonText"),
                        Data = new AdaptiveCardAction
                        {
                            MsteamsCardAction = new CardAction
                            {
                                Type = Constants.MessageBackActionType,
                            },
                            Command = Constants.OkCommand,
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
