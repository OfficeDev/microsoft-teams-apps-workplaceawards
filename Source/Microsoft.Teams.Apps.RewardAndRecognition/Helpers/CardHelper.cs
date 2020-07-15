// <copyright file="CardHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.RewardAndRecognition.Cards;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// Class that handles the card helper methods.
    /// </summary>
    public static class CardHelper
    {
        /// <summary>
        ///  Represents the Configure admin task module height.
        /// </summary>
        private const int ConfigureAdminTaskModuleHeight = 460;

        /// <summary>
        /// Represents the Configure admin task module width.
        /// </summary>
        private const int ConfigureAdminTaskModuleWidth = 600;

        /// <summary>
        ///  Represents the nomination task module height.
        /// </summary>
        private const int NominationTaskModuleHeight = 600;

        /// <summary>
        /// Represents the nomination task module width.
        /// </summary>
        private const int NominationTaskModuleWidth = 700;

        /// <summary>
        /// Represents the error message task module height.
        /// </summary>
        private const int ErrorMessageTaskModuleHeight = 200;

        /// <summary>
        /// Represents the error message task module width.
        /// </summary>
        private const int ErrorMessageTaskModuleWidth = 400;

        /// <summary>
        /// Represents the endorse message task module height.
        /// </summary>
        private const int EndorseMessageTaskModuleHeight = 220;

        /// <summary>
        /// Represents the endorse message task module width.
        /// </summary>
        private const int EndorseMessageTaskModuleWidth = 480;

        /// <summary>
        /// Get nomination task module on messaging extension action response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the application base Uri.</param>
        /// <param name="instrumentationKey">Instrumentation key of the telemetry client.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamId">Team id from where the ME action is called.</param>
        /// <param name="isCycleRunning">Gets the value false if cycle is not running currently.</param>
        /// <returns>Returns task module response.</returns>
        public static MessagingExtensionActionResponse GetNominationTaskModuleBasedOnMessagingExtensionAction(string applicationBasePath, string instrumentationKey, IStringLocalizer<Strings> localizer, string teamId = null, bool isCycleRunning = true)
        {
            // Show error message if award cycle is not active.
            if (!isCycleRunning)
            {
                return new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo()
                        {
                            Card = ValidationMessageCard.GetErrorAdaptiveCard(localizer.GetString("CycleValidationMessage")),
                            Height = ErrorMessageTaskModuleHeight,
                            Width = ErrorMessageTaskModuleWidth,
                            Title = localizer.GetString("NominatePeopleTitle"),
                        },
                    },
                };
            }

            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Url = $"{applicationBasePath}/nominate-awards?telemetry={instrumentationKey}&teamId={teamId}&theme={{theme}}&locale={{locale}}",
                        Height = NominationTaskModuleHeight,
                        Width = NominationTaskModuleWidth,
                        Title = localizer.GetString("NominatePeopleTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get task module response for error message on messaging extension action response.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns task module response.</returns>
        public static MessagingExtensionActionResponse GetTaskModuleInvalidTeamCard(IStringLocalizer<Strings> localizer)
        {
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = ValidationMessageCard.GetErrorAdaptiveCard(localizer.GetString("InvalidTeamText")),
                        Height = ErrorMessageTaskModuleHeight,
                        Width = ErrorMessageTaskModuleWidth,
                        Title = localizer.GetString("NominatePeopleTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get task module response for error message when bot is invoked from 1:1 chat or group chat on messaging extension action response.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <returns>Returns task module response.</returns>
        public static MessagingExtensionActionResponse GetTaskModuleErrorMessageCard(IStringLocalizer<Strings> localizer)
        {
            return new MessagingExtensionActionResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Card = ValidationMessageCard.GetErrorAdaptiveCard(localizer.GetString("MessagingExtensionErrorMessage")),
                        Height = ErrorMessageTaskModuleHeight,
                        Width = ErrorMessageTaskModuleWidth,
                        Title = localizer.GetString("NominatePeopleTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get task endorsement task module response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="nomineeNames">Nominee to name.</param>
        /// <param name="awardName">Award name.</param>
        /// <param name="rewardCycleEndDate">Cycle end date.</param>
        /// <param name="isEndorsementSuccess">Get the endorsement status.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetEndorseTaskModuleResponse(string applicationBasePath, IStringLocalizer<Strings> localizer, string nomineeNames, string awardName, DateTime rewardCycleEndDate, bool isEndorsementSuccess)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = EndorseCard.GetEndorseStatusCard(applicationBasePath, localizer, awardName, nomineeNames, rewardCycleEndDate, isEndorsementSuccess),
                        Height = EndorseMessageTaskModuleHeight,
                        Width = EndorseMessageTaskModuleWidth,
                        Title = localizer.GetString("EndorseTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get task module response for error/validation message.
        /// </summary>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="command">Get the task module command from the user.</param>
        /// <param name="isCycleClosed">Gets the value true if cycle is closed.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetErrorMessageTaskModuleResponse(IStringLocalizer<Strings> localizer, string command, bool isCycleClosed)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = ValidationMessageCard.GetErrorAdaptiveCard(isCycleClosed ? localizer.GetString("CycleClosedMessage") : localizer.GetString("CycleValidationMessage")),
                        Height = ErrorMessageTaskModuleHeight,
                        Width = ErrorMessageTaskModuleWidth,
                        Title = command == Constants.NominateAction ? localizer.GetString("NominatePeopleTitle") : localizer.GetString("EndorseTitle"),
                    },
                },
            };
        }

        /// <summary>
        /// Get nomination task module response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="instrumentationKey">Telemetry instrumentation key.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamId">Team id from where the ME action is called.</param>
        /// <param name="awardId">Award id to fetch the award details.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetNominateTaskModuleResponse(string applicationBasePath, string instrumentationKey, IStringLocalizer<Strings> localizer, string teamId = null, string awardId = null)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo
                    {
                        Url = $"{applicationBasePath}/nominate-awards?telemetry={instrumentationKey}&teamId={teamId}&awardId={awardId}&theme={{theme}}&locale={{locale}}",
                        Height = NominationTaskModuleHeight,
                        Width = NominationTaskModuleWidth,
                        Title = localizer.GetString("NominatePeopleTitle"),
                        FallbackUrl = $"{applicationBasePath}/nominate-awards?telemetry={instrumentationKey}&teamId={teamId}&awardId={awardId}&theme={{theme}}&locale={{locale}}",
                    },
                },
            };
        }

        /// <summary>
        /// Get configure admin task module response.
        /// </summary>
        /// <param name="applicationBasePath">Represents the Application base Uri.</param>
        /// <param name="instrumentationKey">Telemetry instrumentation key.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="teamId">Team id from where the ME action is called.</param>
        /// <param name="updateAdmin">Gets the boolean value based on task module command.</param>
        /// <returns>Returns task module response.</returns>
        public static TaskModuleResponse GetConfigureAdminTaskModuleResponse(string applicationBasePath, string instrumentationKey, IStringLocalizer<Strings> localizer, string teamId = null, bool updateAdmin = true)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = $"{applicationBasePath}/config-admin-page?telemetry={instrumentationKey}&teamId={teamId}&updateAdmin={updateAdmin}&theme={{theme}}&locale={{locale}}",
                        Height = ConfigureAdminTaskModuleHeight,
                        Width = ConfigureAdminTaskModuleWidth,
                        Title = localizer.GetString("ConfigureAdminTitle"),
                        FallbackUrl = $"{applicationBasePath}/config-admin-page?telemetry={instrumentationKey}&teamId={teamId}&updateAdmin={updateAdmin}&theme={{theme}}&locale={{locale}}",
                    },
                },
            };
        }

        /// <summary>
        /// Methods mentions user in respective channel of which they are part after grouping.
        /// </summary>
        /// <param name="mentionToEmails">List of email ID whom to be mentioned.</param>
        /// <param name="userObjectId">Azure active directory object id of current login user.</param>
        /// <param name="teamId">Team id where bot is installed.</param>
        /// <param name="turnContext">Provides context for a turn of a bot.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        /// <param name="logger">Instance to send logs to the application insights service.</param>
        /// <param name="mentionType">Mention activity type.</param>
        /// <param name="cancellationToken">A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns>A task that sends notification in newly created channel and mention its members.</returns>
        internal static async Task<Activity> GetMentionActivityAsync(IEnumerable<string> mentionToEmails, string userObjectId, string teamId, ITurnContext turnContext, IStringLocalizer<Strings> localizer, ILogger logger, MentionActivityType mentionType, CancellationToken cancellationToken)
        {
            List<Entity> entities = new List<Entity>();
            string text = string.Empty;
            string mentionToText = string.Empty;

            try
            {
                var channelMembers = await TeamsInfo.GetTeamMembersAsync(turnContext, teamId, cancellationToken);
                var mentionToMembers = channelMembers.Where(member => mentionToEmails.Contains(member.Email));
                var mentionByMember = channelMembers.Where(member => member.AadObjectId == userObjectId).First();

                foreach (ChannelAccount member in mentionToMembers)
                {
                    Mention mention = new Mention
                    {
                        Mentioned = member,
                        Text = $"<at>{member.Name}</at>",
                    };
                    entities.Add(mention);
                }

                mentionToText = string.Join(", ", mentionToMembers.Select(member => $"<at>{member.Name}</at>"));

                Mention mentionBy = new Mention
                {
                    Mentioned = mentionByMember,
                    Text = $"<at>{mentionByMember.Name}</at>",
                };

                switch (mentionType)
                {
                    case MentionActivityType.SetAdmin:
                        entities.Add(mentionBy);
                        text = localizer.GetString("SetAdminMentionText", mentionToText, mentionBy.Text);
                        break;
                    case MentionActivityType.Nomination:
                        entities.Add(mentionBy);
                        text = localizer.GetString("NominationMentionText", mentionToText, mentionBy.Text);
                        break;
                    case MentionActivityType.Winner:
                        text = $"{localizer.GetString("WinnerMentionText")} {mentionToText}.";
                        break;
                    default:
                        break;
                }

                Activity notificationActivity = MessageFactory.Text(text);
                notificationActivity.Entities = entities;
                return notificationActivity;
            }
#pragma warning disable CA1031 // Catching general exceptions to log exception details in telemetry client.
            catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions to log exception details in telemetry client.
            {
                logger.LogError(ex, $"Error while mentioning channel member in respective channels.");
                return null;
            }
        }
    }
}