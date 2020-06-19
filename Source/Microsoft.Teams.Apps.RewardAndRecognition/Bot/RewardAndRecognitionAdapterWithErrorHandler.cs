// <copyright file="RewardAndRecognitionAdapterWithErrorHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Bot
{
    using System;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// Implements Error Handler.
    /// </summary>
    public class RewardAndRecognitionAdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="RewardAndRecognitionAdapterWithErrorHandler"/> class.
        /// </summary>
        /// <param name="configuration">Application configurations.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="rewardAndRecognitionActivityMiddleware">Represents middle ware that can operate on incoming activities.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="conversationState">conversationState.</param>
        public RewardAndRecognitionAdapterWithErrorHandler(IConfiguration configuration, ILogger<RewardAndRecognitionAdapterWithErrorHandler> logger, RewardAndRecognitionActivityMiddleware rewardAndRecognitionActivityMiddleware, IStringLocalizer<Strings> localizer, ConversationState conversationState = null)
            : base(configuration)
        {
            if (rewardAndRecognitionActivityMiddleware == null)
            {
                throw new ArgumentNullException(nameof(rewardAndRecognitionActivityMiddleware));
            }

            // Add activity middle ware to the adapter's middle ware pipeline
            this.Use(rewardAndRecognitionActivityMiddleware);

            this.OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                logger.LogError(exception, $"Exception caught : {exception.Message}");

                // Send a catch-all apology to the user.
                await turnContext.SendActivityAsync(localizer.GetString("ErrorMessage"));

                if (conversationState != null)
                {
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext);
                    }
#pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client.
                    catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client.
                    {
                        logger.LogError(ex, $"Exception caught on attempting to Delete ConversationState : {ex.Message}");
                    }
                }
            };
        }
    }
}
