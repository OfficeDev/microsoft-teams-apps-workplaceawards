// <copyright file="NominationReminderBackgroundServiceHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Cards;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Helper class which implements <see cref="INominationReminderBackgroundServiceHelper"/>
    /// to create and send nomination reminders.
    /// </summary>
    public class NominationReminderBackgroundServiceHelper : INominationReminderBackgroundServiceHelper
    {
        /// <summary>
        /// Nominate reminder notification days back.
        /// </summary>
        private const int LookBackDays = 3;

        /// <summary>
        /// Retry policy with jitter, retry twice with a jitter delay of up to 1 sec. Retry for HTTP 429(transient error)/502 bad gateway.
        /// </summary>
        /// <remarks>
        /// Reference: https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry#new-jitter-recommendation.
        /// </remarks>
        private static AsyncRetryPolicy retryPolicy = Policy.Handle<ErrorResponseException>(
            ex => ex.Response.StatusCode == HttpStatusCode.TooManyRequests || ex.Response.StatusCode == HttpStatusCode.BadGateway)
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 2));

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Microsoft application credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Represents a set of key/value application configuration properties.
        /// </summary>
        private readonly IOptions<RewardAndRecognitionActivityHandlerOptions> options;

        /// <summary>
        /// Helper for storing reward cycle details to azure table storage.
        /// </summary>
        private readonly IRewardCycleStorageProvider rewardCycleStorageProvider;

        /// <summary>
        /// Helper for fetching reward details from azure table storage.
        /// </summary>
        private readonly IAwardsStorageProvider awardsStorageProvider;

        /// <summary>
        /// Helper for fetching teams details from azure table storage.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// Provider to store logs in Azure Application Insights.
        /// </summary>
        private readonly ILogger<NominationReminderBackgroundServiceHelper> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="NominationReminderBackgroundServiceHelper"/> class.
        /// </summary>
        /// <param name="rewardCycleStorageProvider">Reward cycle storage provider.</param>
        /// <param name="awardsStorageProvider">Award storage provider.</param>
        /// <param name="teamStorageProvider">Teams storage provider.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="options">A set of key/value application configuration properties.</param>
        /// <param name="adapter">Bot adapter.</param>
        /// <param name="microsoftAppCredentials">MicrosoftAppCredentials of bot.</param>
        public NominationReminderBackgroundServiceHelper(
            IRewardCycleStorageProvider rewardCycleStorageProvider,
            ITeamStorageProvider teamStorageProvider,
            IAwardsStorageProvider awardsStorageProvider,
            ILogger<NominationReminderBackgroundServiceHelper> logger,
            IStringLocalizer<Strings> localizer,
            IOptions<RewardAndRecognitionActivityHandlerOptions> options,
            IBotFrameworkHttpAdapter adapter,
            MicrosoftAppCredentials microsoftAppCredentials)
        {
            this.rewardCycleStorageProvider = rewardCycleStorageProvider;
            this.logger = logger;
            this.localizer = localizer;
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.adapter = adapter;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.awardsStorageProvider = awardsStorageProvider;
            this.teamStorageProvider = teamStorageProvider;
        }

        /// <summary>
        /// This method is used to send nomination reminder notification.
        /// </summary>
        /// <returns>Returns true if nomination reminder sent successfully else false.</returns>
        public async Task<bool> SendNominationReminderNotificationAsync()
        {
            var activeRewardCycle = await this.rewardCycleStorageProvider.GetActiveRewardCycleForAllTeamsAsync();
            foreach (var currentCycle in activeRewardCycle)
            {
                try
                {
                    if (currentCycle.RewardCycleEndDate.ToUniversalTime().Day == DateTime.UtcNow.AddDays(-LookBackDays).Day)
                    {
                        // Send nomination reminder notification
                        await this.SendCardToTeamAsync(currentCycle);
                    }
                }
#pragma warning disable CA1031 // Catching general exceptions to unblock iteration for sending reminder notifications to next team.
                catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to unblock iteration for sending reminder notifications to next team.
                {
                    this.logger.LogError(ex, $"Error occurred while sending reminder notification for team: {currentCycle.TeamId}.");
                }
            }

            return true;
        }

        /// <summary>
        /// Send the nomination reminder notification to specified team.
        /// </summary>
        /// <param name="rewardCycleEntity">Reward cycle model object.</param>
        /// <returns>A task that sends notification card in channel.</returns>
        private async Task SendCardToTeamAsync(RewardCycleEntity rewardCycleEntity)
        {
            rewardCycleEntity = rewardCycleEntity ?? throw new ArgumentNullException(nameof(rewardCycleEntity));

            var awardsList = await this.awardsStorageProvider.GetAwardsAsync(rewardCycleEntity.TeamId);
            var valuesFromTaskModule = new TaskModuleResponseDetails()
            {
                RewardCycleStartDate = rewardCycleEntity.RewardCycleStartDate,
                RewardCycleEndDate = rewardCycleEntity.RewardCycleEndDate,
                RewardCycleId = rewardCycleEntity.CycleId,
            };

            var teamDetails = await this.teamStorageProvider.GetTeamDetailAsync(rewardCycleEntity.TeamId);
            string serviceUrl = teamDetails.ServiceUrl;

            MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
            string teamGeneralChannelId = rewardCycleEntity.TeamId;

            this.logger.LogInformation($"sending notification to channel id - {teamGeneralChannelId}");

            await retryPolicy.ExecuteAsync(async () =>
            {
                try
                {
                    var conversationParameters = new ConversationParameters()
                    {
                        ChannelData = new TeamsChannelData() { Channel = new ChannelInfo() { Id = rewardCycleEntity.TeamId } },
                        Activity = (Activity)MessageFactory.Carousel(NominateCarouselCard.GetAwardNominationCards(this.options.Value.AppBaseUri, awardsList, this.localizer, valuesFromTaskModule)),
                    };

                    Activity mentionActivity = MessageFactory.Text(this.localizer.GetString("NominationReminderNotificationText"));

                    await ((BotFrameworkAdapter)this.adapter).CreateConversationAsync(
                        Constants.TeamsBotFrameworkChannelId,
                        serviceUrl,
                        this.microsoftAppCredentials,
                        conversationParameters,
                        async (conversationTurnContext, conversationCancellationToken) =>
                        {
                            await conversationTurnContext.SendActivityAsync(mentionActivity, conversationCancellationToken);
                        },
                        default);
                }
                catch (Exception ex)
                {
                    this.logger.LogError(ex, "Error while sending mention card notification to channel.");
                    throw;
                }
            });
        }
    }
}
