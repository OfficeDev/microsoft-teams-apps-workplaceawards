// <copyright file="NotificationController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Authentication.AuthenticationPolicy;
    using Microsoft.Teams.Apps.RewardAndRecognition.Cards;
    using Microsoft.Teams.Apps.RewardAndRecognition.Helpers;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;

    /// <summary>
    /// This ASP controller is created to handle award requests and leverages TeamMemberUserPolicy for authorization.
    /// Dependency injection will provide the storage implementation and logger.
    /// Inherits <see cref="BaseRewardAndRecognitionController"/> to gather user claims for all incoming requests.
    /// The class provides endpoint to send proactive notifications to required audience.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    [Authorize(PolicyNames.MustBeTeamMemberUserPolicy)]
    public class NotificationController : BaseRewardAndRecognitionController
    {
        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly IBotFrameworkHttpAdapter adapter;

        /// <summary>
        /// Represents a set of key/value settings properties.
        /// </summary>
        private readonly IOptions<RewardAndRecognitionActivityHandlerOptions> botSettingsOptions;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<NotificationController> logger;

        /// <summary>
        /// Provider for fetching information about team details from storage table.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Microsoft application credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationController"/> class.
        /// </summary>
        /// <param name="adapter">bot adapter.</param>
        /// <param name="botSettingsOptions">Bot settings options.</param>
        /// <param name="logger">Provider to store logs in Azure Application Insights.</param>
        /// <param name="teamStorageProvider">Store or update teams details in Azure table storage.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="microsoftAppCredentials">MicrosoftAppCredentials of bot.</param>
        public NotificationController(IBotFrameworkHttpAdapter adapter, IOptions<RewardAndRecognitionActivityHandlerOptions> botSettingsOptions, ILogger<NotificationController> logger, ITeamStorageProvider teamStorageProvider, IStringLocalizer<Strings> localizer, MicrosoftAppCredentials microsoftAppCredentials)
        {
            this.adapter = adapter;
            this.botSettingsOptions = botSettingsOptions;
            this.logger = logger;
            this.teamStorageProvider = teamStorageProvider;
            this.localizer = localizer;
            this.microsoftAppCredentials = microsoftAppCredentials;
        }

        /// <summary>
        /// Send proactive notification card in channel
        /// for all award winners.
        /// </summary>
        /// <param name="details">Notification details.</param>
        /// <returns>Sends winner card to teams channel.</returns>
        [HttpPost("winnernotification")]
        [Authorize(PolicyNames.MustBeTeamCaptainUserPolicy)]
        public async Task<IActionResult> WinnerNominationAsync([FromBody]AwardWinner details)
        {
            try
            {
                if (details == null || details.Winners == null)
                {
                    return this.BadRequest(new { message = "Award winner details can not be null." });
                }

                var emails = string.Join(",", details.Winners.Select(row => row.NomineeUserPrincipalNames)).Split(",").Select(row => row.Trim()).Distinct();
                string teamId = details.TeamId;
                var claims = this.GetUserClaims();
                var teamDetails = await this.teamStorageProvider.GetTeamDetailAsync(teamId);
                string serviceUrl = teamDetails.ServiceUrl;
                string appBaseUrl = this.botSettingsOptions.Value.AppBaseUri;
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);
                var conversationParameters = new ConversationParameters()
                {
                    ChannelData = new TeamsChannelData() { Channel = new ChannelInfo() { Id = teamId } },
                    Activity = (Activity)MessageFactory.Carousel(WinnerCarouselCard.GetAwardWinnerCard(appBaseUrl, details.Winners, this.localizer)),
                    Bot = new ChannelAccount() { Id = this.microsoftAppCredentials.MicrosoftAppId },
                    IsGroup = true,
                    TenantId = this.botSettingsOptions.Value.TenantId,
                };

                await ((BotFrameworkAdapter)this.adapter).CreateConversationAsync(
                    Constants.TeamsBotFrameworkChannelId,
                    serviceUrl,
                    this.microsoftAppCredentials,
                    conversationParameters,
                    async (turnContext, cancellationToken) =>
                    {
                        Activity mentionActivity = await CardHelper.GetMentionActivityAsync(emails, claims.FromId, teamId, turnContext, this.localizer, this.logger, MentionActivityType.Winner, default);
                        await turnContext.SendActivityAsync(mentionActivity, cancellationToken);
                        await turnContext.SendActivityAsync(MessageFactory.Text(this.localizer.GetString("ViewWinnerTabText")), cancellationToken);
                    }, default);

                // Let the caller know proactive messages have been sent
                return this.Ok(true);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, "problem while sending the winner card.");
                throw;
            }
        }
    }
}
