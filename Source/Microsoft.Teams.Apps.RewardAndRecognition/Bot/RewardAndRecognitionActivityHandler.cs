// <copyright file="RewardAndRecognitionActivityHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition.Cards;
    using Microsoft.Teams.Apps.RewardAndRecognition.Helpers;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;
    using Microsoft.Teams.Apps.RewardAndRecognition.Providers;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// The RewardAndRecognitionActivityHandler is responsible for reacting to incoming events from Teams sent from BotFramework.
    /// </summary>
    public sealed class RewardAndRecognitionActivityHandler : TeamsActivityHandler
    {
        /// <summary>
        /// Represents the conversation type as channel.
        /// </summary>
        private const string ChannelConversationType = "channel";

        /// <summary>
        /// Represents the conversation type as personal.
        /// </summary>
        private const string PersonalConversationType = "personal";

        /// <summary>
        /// A set of key/value application configuration properties for Activity settings.
        /// </summary>
        private readonly IOptions<RewardAndRecognitionActivityHandlerOptions> options;

        /// <summary>
        /// Instrumentation key of the telemetry client.
        /// </summary>
        private readonly string instrumentationKey;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<RewardAndRecognitionActivityHandler> logger;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Instance of Application Insights Telemetry client.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Microsoft application credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Bot adapter.
        /// </summary>
        private readonly BotFrameworkAdapter botAdapter;

        /// <summary>
        /// Provider for fetching information about admin details from storage table.
        /// </summary>
        private readonly IConfigureAdminStorageProvider configureAdminStorageProvider;

        /// <summary>
        /// Provider for fetching information about team details from storage table.
        /// </summary>
        private readonly ITeamStorageProvider teamStorageProvider;

        /// <summary>
        /// Provider for fetching information about awards from storage table.
        /// </summary>
        private readonly IAwardsStorageProvider awardsStorageProvider;

        /// <summary>
        /// Provider for fetching information about endorsement details from storage table.
        /// </summary>
        private readonly IEndorsementsStorageProvider endorseDetailStorageProvider;

        /// <summary>
        /// Provider for fetching information about active award cycle details from storage table.
        /// </summary>
        private readonly IRewardCycleStorageProvider rewardCycleStorageProvider;

        /// <summary>
        /// Provider to search nomination details in Azure search service.
        /// </summary>
        private readonly IAwardNominationSearchService nominateDetailSearchService;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardAndRecognitionActivityHandler"/> class.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="telemetryClient">The application insights telemetry client. </param>
        /// <param name="options">The options.</param>
        /// <param name="telemetryOptions">Telemetry instrumentation key</param>
        /// <param name="configureAdminStorageProvider">Provider for fetching information about admin details from storage table.</param>
        /// <param name="teamStorageProvider">Provider for fetching information about team details from storage table.</param>
        /// <param name="awardsStorageProvider">Provider for fetching information about awards from storage table.</param>
        /// <param name="endorseDetailStorageProvider">Provider for fetching information about endorsement details from storage table.</param>
        /// <param name="rewardCycleStorageProvider">Provider for fetching information about active award cycle details from storage table.</param>
        /// <param name="nominateDetailSearchService">Provider to search nomination details in Azure search service.</param>
        /// <param name="botAdapter">Bot adapter.</param>
        /// <param name="microsoftAppCredentials">MicrosoftAppCredentials of bot.</param>
        public RewardAndRecognitionActivityHandler(
            ILogger<RewardAndRecognitionActivityHandler> logger,
            IStringLocalizer<Strings> localizer,
            TelemetryClient telemetryClient,
            IOptions<RewardAndRecognitionActivityHandlerOptions> options,
            IOptions<TelemetryOptions> telemetryOptions,
            IConfigureAdminStorageProvider configureAdminStorageProvider,
            ITeamStorageProvider teamStorageProvider,
            IAwardsStorageProvider awardsStorageProvider,
            IEndorsementsStorageProvider endorseDetailStorageProvider,
            IRewardCycleStorageProvider rewardCycleStorageProvider,
            IAwardNominationSearchService nominateDetailSearchService,
            BotFrameworkAdapter botAdapter,
            MicrosoftAppCredentials microsoftAppCredentials)
        {
            options = options ?? throw new ArgumentNullException(nameof(options));
            telemetryOptions = telemetryOptions ?? throw new ArgumentNullException(nameof(telemetryOptions));

            this.logger = logger;
            this.localizer = localizer;
            this.telemetryClient = telemetryClient;
            this.options = options;
            this.instrumentationKey = telemetryOptions.Value.InstrumentationKey;
            this.configureAdminStorageProvider = configureAdminStorageProvider;
            this.teamStorageProvider = teamStorageProvider;
            this.awardsStorageProvider = awardsStorageProvider;
            this.endorseDetailStorageProvider = endorseDetailStorageProvider;
            this.rewardCycleStorageProvider = rewardCycleStorageProvider;
            this.nominateDetailSearchService = nominateDetailSearchService;
            this.botAdapter = botAdapter;
            this.microsoftAppCredentials = microsoftAppCredentials;
        }

        /// <summary>
        /// Handle when a message is addressed to the bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// For more information on bot messaging in Teams, see the documentation
        /// https://docs.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-basics?tabs=dotnet#receive-a-message .
        /// </remarks>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
                this.RecordEvent(nameof(this.OnMessageActivityAsync), turnContext);
                await this.SendTypingIndicatorAsync(turnContext);
                await turnContext.SendActivityAsync(MessageFactory.Text(this.localizer.GetString("UnsupportedBotCommand")), cancellationToken);
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error processing message: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Overriding to send welcome card once Bot/ME is installed in team.
        /// </summary>
        /// <param name="membersAdded">A list of all the members added to the conversation, as described by the conversation update activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Welcome card  when bot is added first time by user.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            var activity = turnContext.Activity;
            this.logger.LogInformation($"conversationType: {activity.Conversation?.ConversationType}, membersAdded: {membersAdded?.Count}");

            if (membersAdded.Any(member => member.Id == activity.Recipient.Id) && activity.Conversation.ConversationType == ChannelConversationType)
            {
                this.logger.LogInformation($"Bot added {activity.Conversation.Id}");

                // Storing team information to storage
                var teamsDetails = activity.TeamsGetTeamInfo();
                TeamEntity teamEntity = new TeamEntity
                {
                    TeamId = teamsDetails.Id,
                    BotInstalledOn = DateTime.UtcNow,
                    ServiceUrl = turnContext.Activity.ServiceUrl,
                    RowKey = teamsDetails.Id,
                };

                bool operationStatus = await this.teamStorageProvider.StoreOrUpdateTeamDetailAsync(teamEntity);
                if (!operationStatus)
                {
                    this.logger.LogInformation($"Unable to store bot Installation detail in table storage.");
                }

                await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.GetCard(this.options.Value.AppBaseUri, this.localizer)), cancellationToken);
            }
        }

        /// <summary>
        /// Overriding to send card when award admin member is removed from team.
        /// </summary>
        /// <param name="membersRemoved">A member removed from team, as described by the conversation update activity.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Notification card  when bot or member is removed from team.</returns>
        protected override async Task OnMembersRemovedAsync(IList<ChannelAccount> membersRemoved, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            var activity = turnContext.Activity;
            this.logger.LogInformation($"conversationType: {activity.Conversation?.ConversationType}, membersRemoved: {membersRemoved?.Count}");
            if (activity.Conversation.ConversationType == ChannelConversationType)
            {
                var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
                var admin = await this.configureAdminStorageProvider.GetAdminDetailAsync(teamsDetails.Id);

                // In application, there is a persona named 'Champion' who is the only person in team to create reward cycle, add awards and publish results.
                // In case if the Champion is removed from team, then the bot sends a new card to set up new Champion.
                if (membersRemoved.Any(member => member.AadObjectId == admin.AdminObjectId))
                {
                    this.logger.LogInformation($"Award captain is removed from team. {activity.Conversation.Id}");
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(WelcomeCard.ConfigureNewAdminCard(this.localizer)), cancellationToken);
                }

                // Deleting team information from storage when bot is uninstalled from a team.
                else if (membersRemoved.Any(member => member.Id == activity.Recipient.Id))
                {
                    this.logger.LogInformation($"Bot removed {activity.Conversation.Id}");
                    var teamEntity = await this.teamStorageProvider.GetTeamDetailAsync(teamsDetails.Id);
                    bool operationStatus = await this.teamStorageProvider.DeleteTeamDetailAsync(teamEntity);
                    if (!operationStatus)
                    {
                        this.logger.LogInformation($"Unable to remove team details from table storage.");
                    }
                }
            }
        }

        /// <summary>
        /// Handle message extension action fetch task received by the bot.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="action">Messaging extension action value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionfetchtaskasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            if (turnContext.Activity?.Conversation?.ConversationType == PersonalConversationType)
            {
                return CardHelper.GetTaskModuleErrorMessageCard(this.localizer);
            }

            if (!await this.CheckTeamsValidationAsync(turnContext, cancellationToken))
            {
                return CardHelper.GetTaskModuleInvalidTeamCard(this.localizer);
            }

            this.RecordEvent(nameof(this.OnTeamsMessagingExtensionFetchTaskAsync), turnContext);

            var activity = turnContext.Activity;
            var teamDetails = turnContext.Activity.TeamsGetTeamInfo();
            var rewardCycleDetail = await this.rewardCycleStorageProvider.GetCurrentRewardCycleAsync(teamDetails.Id);
            bool isCycleRunning = !(rewardCycleDetail == null || rewardCycleDetail.RewardCycleState == (int)RewardCycleState.Inactive);

            return CardHelper.GetNominationTaskModuleBasedOnMessagingExtensionAction(this.options.Value.AppBaseUri, this.instrumentationKey, this.localizer, teamDetails.Id, isCycleRunning);
        }

        /// <summary>
        /// Invoked when the user submits a response from messaging extension.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="action">Messaging extension action commands.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionsubmitactionasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            action = action ?? throw new ArgumentNullException(nameof(action));

            this.RecordEvent(nameof(this.OnTeamsMessagingExtensionSubmitActionAsync), turnContext);
            var valuesFromTaskModule = JsonConvert.DeserializeObject<TaskModuleResponseDetails>(action.Data.ToString());
            if (valuesFromTaskModule.Command.ToUpperInvariant() == Constants.SaveNominatedDetailsAction)
            {
                var mentionActivity = await CardHelper.GetMentionActivityAsync(
                    valuesFromTaskModule.NomineeUserPrincipalNames.Split(",").Select(row => row.Trim()).ToList(),
                    turnContext.Activity.From.AadObjectId,
                    valuesFromTaskModule.TeamId,
                    turnContext,
                    this.localizer,
                    this.logger,
                    MentionActivityType.Nomination,
                    cancellationToken);

                var notificationCard = EndorseCard.GetEndorseCard(this.options.Value.AppBaseUri, this.localizer, valuesFromTaskModule);

                await this.SendCardAndMentionsAsync(turnContext, notificationCard, mentionActivity);
                this.logger.LogInformation("Award nomination card sent successfully.");

                return null;
            }

            this.logger.LogWarning($"Unsupported bot command: {valuesFromTaskModule.Command}");

            return null;
        }

        /// <summary>
        /// When OnTurn method receives a fetch invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            var activity = (Activity)turnContext.Activity;
            this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);

            var teamsDetails = activity.TeamsGetTeamInfo();
            var valuesforTaskModule = JsonConvert.DeserializeObject<AdaptiveCardAction>(((JObject)activity.Value).GetValue("data", StringComparison.Ordinal)?.ToString());
            var rewardCycleDetail = await this.rewardCycleStorageProvider.GetCurrentRewardCycleAsync(teamsDetails.Id);
            if (rewardCycleDetail != null && rewardCycleDetail.CycleId != valuesforTaskModule.RewardCycleId && valuesforTaskModule.Command != Constants.ConfigureAdminAction)
            {
                return CardHelper.GetErrorMessageTaskModuleResponse(localizer: this.localizer, command: valuesforTaskModule.Command, isCycleClosed: true);
            }
            else if ((rewardCycleDetail == null || rewardCycleDetail.RewardCycleState == (int)RewardCycleState.Inactive) && valuesforTaskModule.Command != Constants.ConfigureAdminAction)
            {
                return CardHelper.GetErrorMessageTaskModuleResponse(localizer: this.localizer, command: valuesforTaskModule.Command, isCycleClosed: false);
            }

            switch (valuesforTaskModule.Command)
            {
                // Fetch task module to show configure admin card
                case Constants.ConfigureAdminAction:
                    this.logger.LogInformation("Fetch task module to show configure admin card.");
                    return CardHelper.GetConfigureAdminTaskModuleResponse(this.options.Value.AppBaseUri, this.instrumentationKey, this.localizer, teamsDetails.Id, updateAdmin: false);

                // Fetch and show task module to endorse an award nomination
                case Constants.EndorseAction:
                    bool isEndorsementSuccess = await this.CheckEndorseStatusAsync(turnContext, valuesforTaskModule, cancellationToken);
                    this.logger.LogInformation("Fetch and show task module to endorse an award nomination.");
                    return CardHelper.GetEndorseTaskModuleResponse(applicationBasePath: this.options.Value.AppBaseUri, this.localizer, valuesforTaskModule.NomineeNames, valuesforTaskModule.AwardName, rewardCycleDetail.RewardCycleEndDate, isEndorsementSuccess);

                // Fetch and show task module to show new nominate card.
                case Constants.NominateAction:
                    this.logger.LogInformation("Fetch and show task module to configure new nominate award card.");
                    return CardHelper.GetNominateTaskModuleResponse(this.options.Value.AppBaseUri, this.instrumentationKey, this.localizer, teamsDetails.Id, valuesforTaskModule.AwardId);

                default:
                    this.logger.LogInformation($"Invalid command for task module fetch activity.Command is : {valuesforTaskModule.Command} ");
                    await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                    return null;
            }
        }

        /// <summary>
        /// When OnTurn method receives a submit invoke activity on bot turn, it calls this method.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that represents a task module response.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            try
            {
                turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

                var activity = (Activity)turnContext.Activity;
                this.RecordEvent(nameof(this.OnTeamsTaskModuleFetchAsync), turnContext);
                Activity mentionActivity;
                var valuesFromTaskModule = JsonConvert.DeserializeObject<TaskModuleResponseDetails>(((JObject)activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase)?.ToString());
                switch (valuesFromTaskModule.Command.ToUpperInvariant())
                {
                    // Command to send award admin card on save admin action.
                    case Constants.SaveAdminDetailsAction:
                        mentionActivity = await CardHelper.GetMentionActivityAsync(
                            valuesFromTaskModule.AdminUserPrincipalName.Split(",").ToList(),
                            turnContext.Activity.From.AadObjectId,
                            valuesFromTaskModule.TeamId,
                            turnContext,
                            this.localizer,
                            this.logger,
                            MentionActivityType.SetAdmin,
                            cancellationToken);
                        var cardDetail = AdminCard.GetAdminCard(this.localizer, valuesFromTaskModule);
                        await this.SendCardAndMentionsAsync(turnContext, cardDetail, mentionActivity);
                        this.logger.LogInformation("Admin has been configured successfully.");

                        break;

                    // Command to update award admin card
                    case Constants.UpdateAdminDetailCommand:
                        mentionActivity = await CardHelper.GetMentionActivityAsync(
                            valuesFromTaskModule.AdminUserPrincipalName.Split(",").ToList(),
                            turnContext.Activity.From.AadObjectId,
                            valuesFromTaskModule.TeamId,
                            turnContext,
                            this.localizer,
                            this.logger,
                            MentionActivityType.SetAdmin,
                            cancellationToken);

                        var notificationCard = (Activity)MessageFactory.Attachment(AdminCard.GetAdminCard(this.localizer, valuesFromTaskModule));

                        // Split here extracts the activity id from turn context conversation
                        notificationCard.Id = turnContext.Activity.Conversation.Id.Split(';')[1].Split("=")[1];
                        notificationCard.Conversation = turnContext.Activity.Conversation;
                        await turnContext.UpdateActivityAsync(notificationCard);
                        await turnContext.SendActivityAsync(mentionActivity);
                        this.logger.LogInformation("Admin card is updated successfully.");
                        break;

                    // Command to show list of awards ready for nomination
                    case Constants.NominateAction:
                        var awardsList = await this.awardsStorageProvider.GetAwardsAsync(valuesFromTaskModule.TeamId);
                        await turnContext.SendActivityAsync(MessageFactory.Carousel(NominateCarouselCard.GetAwardNominationCards(this.options.Value.AppBaseUri, awardsList, this.localizer, valuesFromTaskModule)));
                        this.logger.LogInformation("Nomination carousel card is sent successfully.");
                        break;

                    // Command to save nominated user details
                    case Constants.SaveNominatedDetailsAction:
                        turnContext.Activity.Conversation.Id = valuesFromTaskModule.TeamId;
                        var endorsementCard = EndorseCard.GetEndorseCard(this.options.Value.AppBaseUri, this.localizer, valuesFromTaskModule);
                        mentionActivity = await CardHelper.GetMentionActivityAsync(
                            valuesFromTaskModule.NomineeUserPrincipalNames.Split(",").Select(row => row.Trim()).ToList(),
                            turnContext.Activity.From.AadObjectId,
                            valuesFromTaskModule.TeamId,
                            turnContext,
                            this.localizer,
                            this.logger,
                            MentionActivityType.Nomination,
                            cancellationToken);

                        await this.SendCardAndMentionsAsync(turnContext, endorsementCard, mentionActivity);
                        this.logger.LogInformation("Award nomination for user sent successfully");
                        break;

                    // Commands to close task modules
                    case Constants.OkCommand:
                    case Constants.CancelCommand:
                        this.logger.LogInformation($"{valuesFromTaskModule.Command.ToUpperInvariant()} is called. [note] - no actions are performed.");
                        break;

                    default:
                        this.logger.LogInformation($"Invalid command for task module submit activity.Command is : {valuesFromTaskModule.Command} ");
                        await turnContext.SendActivityAsync(this.localizer.GetString("ErrorMessage"));
                        break;
                }

                return null;
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Error at OnTeamsTaskModuleSubmitAsync(): {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Invoked when the user opens the messaging extension or searching any content in it.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Messaging extension response object to fill compose extension section.</returns>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionqueryasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionQuery query,
            CancellationToken cancellationToken)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));

            IInvokeActivity turnContextActivity = turnContext.Activity;
            try
            {
                if (turnContextActivity != null && (turnContextActivity.Conversation?.ConversationType == null || turnContextActivity.Conversation?.ConversationType == PersonalConversationType))
                {
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = new MessagingExtensionResult
                        {
                            Text = this.localizer.GetString("MessagingExtensionErrorMessage"),
                            Type = "message",
                        },
                    };
                }

                if (!await this.CheckTeamsValidationAsync(turnContext, cancellationToken))
                {
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = new MessagingExtensionResult
                        {
                            Text = this.localizer.GetString("InvalidTeamText"),
                            Type = "message",
                        },
                    };
                }

                MessagingExtensionQuery messageExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(turnContextActivity.Value.ToString());
                string searchQuery = SearchHelper.GetSearchQueryString(messageExtensionQuery);
                turnContextActivity.TryGetChannelData<TeamsChannelData>(out var teamsChannelData);
                var cycleStatus = await this.rewardCycleStorageProvider.GetCurrentRewardCycleAsync(teamsChannelData.Team.Id);

                if (cycleStatus != null)
                {
                    return new MessagingExtensionResponse
                    {
                        ComposeExtension = await SearchHelper.GetSearchResultAsync(
                            this.options.Value.AppBaseUri,
                            searchQuery,
                            cycleStatus.CycleId,
                            teamsChannelData.Team.Id,
                            messageExtensionQuery.QueryOptions.Count,
                            messageExtensionQuery.QueryOptions.Skip,
                            this.nominateDetailSearchService,
                            this.localizer),
                    };
                }

                return new MessagingExtensionResponse
                {
                    ComposeExtension = new MessagingExtensionResult
                    {
                        Text = this.localizer.GetString("CycleValidationMessage"),
                        Type = "message",
                    },
                };
            }
            catch (Exception ex)
            {
                this.logger.LogError(ex, $"Failed to handle the messaging extension command {turnContextActivity.Name}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Validates endorsement status.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="valuesforTaskModule">Get the binded values from the card.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Returns the true, if endorsement is successful, else false.</returns>
        private async Task<bool> CheckEndorseStatusAsync(ITurnContext<IInvokeActivity> turnContext, AdaptiveCardAction valuesforTaskModule, CancellationToken cancellationToken)
        {
            var teamsDetails = turnContext.Activity.TeamsGetTeamInfo();
            var teamsChannelAccounts = await TeamsInfo.GetTeamMembersAsync(turnContext, teamsDetails.Id, cancellationToken);
            var userDetails = teamsChannelAccounts.Where(member => member.AadObjectId == turnContext.Activity.From.AadObjectId).FirstOrDefault();
            var endorseEntity = await this.endorseDetailStorageProvider.GetEndorsementsAsync(teamsDetails.Id, valuesforTaskModule.RewardCycleId, valuesforTaskModule.NomineeObjectIds);
            var result = endorseEntity.Where(row => row.EndorsedForAwardId == valuesforTaskModule.AwardId && row.EndorsedByObjectId == userDetails.AadObjectId).FirstOrDefault();
            if (result == null)
            {
                var endorsedetails = new EndorsementEntity
                {
                    TeamId = teamsDetails.Id,
                    EndorsedByObjectId = userDetails.AadObjectId,
                    EndorsedByUserPrincipalName = userDetails.UserPrincipalName,
                    EndorsedForAward = valuesforTaskModule.AwardName,
                    EndorseeUserPrincipalName = valuesforTaskModule.NomineeUserPrincipalNames,
                    EndorseeObjectId = valuesforTaskModule.NomineeObjectIds,
                    EndorsedOn = DateTime.UtcNow,
                    EndorsedForAwardId = valuesforTaskModule.AwardId,
                    AwardCycle = valuesforTaskModule.RewardCycleId,
                };

                return await this.endorseDetailStorageProvider.StoreOrUpdateEndorsementDetailAsync(endorsedetails);
            }

            return false;
        }

        /// <summary>
        /// Validates Teams metadata is present in Azure Table Storage.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>Returns the true, if teams metadata is present in Azure Table Storage, else false.</returns>
        private async Task<bool> CheckTeamsValidationAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            var teamInformation = turnContext.Activity.TeamsGetTeamInfo();
            if (teamInformation == null)
            {
                this.logger.LogInformation($"Validation failed:  Teams metadata is not present in Azure Table Storage.");
                return false;
            }

            try
            {
                await TeamsInfo.GetTeamMembersAsync(turnContext, teamInformation.Id, cancellationToken);
                return true;
            }
#pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client.
            catch
#pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client.
            {
                this.logger.LogInformation($"Validation failed:  Bot has not been added for this team.");
                return false;
            }
        }

        /// <summary>
        /// Records event occurred in the application in Application Insights telemetry client.
        /// </summary>
        /// <param name="eventName"> Name of the event.</param>
        /// <param name="turnContext"> Context object containing information cached for a single turn of conversation with a user.</param>
        private void RecordEvent(string eventName, ITurnContext turnContext)
        {
            this.telemetryClient.TrackEvent(eventName, new Dictionary<string, string>
            {
                { "userId", turnContext.Activity.From.AadObjectId },
                { "tenantId", turnContext.Activity.Conversation.TenantId },
                { "teamId", turnContext.Activity.Conversation.Id },
                { "channelId", turnContext.Activity.ChannelId },
            });
        }

        /// <summary>
        /// Send typing indicator to the user.
        /// </summary>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <returns>A task that represents typing indicator activity.</returns>
        private async Task SendTypingIndicatorAsync(ITurnContext turnContext)
        {
            try
            {
                var typingActivity = turnContext.Activity.CreateReply();
                typingActivity.Type = ActivityTypes.Typing;
                await turnContext.SendActivityAsync(typingActivity);
            }
#pragma warning disable CA1031 // Catching general exceptions that might arise during send typing indicator activity to user, to avoid blocking next execution
            catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions that might arise during send typing indicator activity to user, to avoid blocking next execution
            {
                // Do not fail on errors sending the typing indicator
                this.logger.LogWarning(ex, "Failed to send a typing indicator.");
            }
        }

        /// <summary>
        /// Send adaptive card with mentioned activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="mainCard">Adaptive card.</param>
        /// <param name="mentionActivity">Mentioned card activity.</param>
        private async Task SendCardAndMentionsAsync(ITurnContext<IInvokeActivity> turnContext, Attachment mainCard, Activity mentionActivity)
        {
            if (turnContext.Activity.Conversation.ConversationType == ChannelConversationType)
            {
                var channelData = turnContext.Activity.GetChannelData<TeamsChannelData>();
                var conversationParameters = new ConversationParameters()
                {
                    ChannelData = channelData,
                    Activity = (Activity)MessageFactory.Attachment(mainCard),
                };

                await this.botAdapter.CreateConversationAsync(
                    Constants.TeamsBotFrameworkChannelId,
                    turnContext.Activity.ServiceUrl,
                    this.microsoftAppCredentials,
                    conversationParameters,
                    async (conversationTurnContext, conversationCancellationToken) =>
                    {
                        await conversationTurnContext.SendActivityAsync(mentionActivity, conversationCancellationToken);
                    },
                    default);
            }
        }
    }
}