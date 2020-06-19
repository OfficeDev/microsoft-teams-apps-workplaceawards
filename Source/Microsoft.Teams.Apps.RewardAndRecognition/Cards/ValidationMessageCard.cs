// <copyright file="ValidationMessageCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.RewardAndRecognition.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;

    /// <summary>
    /// This class is to render the adaptive card with error message.
    /// </summary>
    public static class ValidationMessageCard
    {
        /// <summary>
        /// Construct the card to render error message to task module.
        /// </summary>
        /// <param name="message">Message to show as error.</param>
        /// <returns>Card attachment.</returns>
        public static Attachment GetErrorAdaptiveCard(string message)
        {
            AdaptiveCard card = new AdaptiveCard(new AdaptiveSchemaVersion(Constants.AdaptiveCardVersion))
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = message,
                        Wrap = true,
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
