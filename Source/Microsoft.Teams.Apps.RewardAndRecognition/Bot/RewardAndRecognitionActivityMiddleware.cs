// <copyright file="RewardAndRecognitionActivityMiddleware.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Bot
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.RewardAndRecognition;

    /// <summary>
    /// Represents middle ware that can operate on incoming activities.
    /// </summary>
    public class RewardAndRecognitionActivityMiddleware : IMiddleware
    {
        /// <summary>
        /// Tenant id of Microsoft Teams where application is installed.
        /// </summary>
        private readonly string tenantId;

        /// <summary>
        /// The current cultures' string localizer.
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Remote Support bot.
        /// </summary>
        private readonly IOptions<RewardAndRecognitionActivityHandlerOptions> options;

        /// <summary>
        /// Provider to store logs in Azure Application Insights.
        /// </summary>
        private readonly ILogger<RewardAndRecognitionActivityMiddleware> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardAndRecognitionActivityMiddleware"/> class.
        /// </summary>
        /// <param name="options"> A set of key/value application configuration properties.</param>
        /// <param name="logger">Provider to store logs in Azure Application Insights.</param>
        /// <param name="localizer">The current cultures' string localizer.</param>
        public RewardAndRecognitionActivityMiddleware(IOptions<RewardAndRecognitionActivityHandlerOptions> options, ILogger<RewardAndRecognitionActivityMiddleware> logger, IStringLocalizer<Strings> localizer)
        {
            this.options = options ?? throw new ArgumentNullException(nameof(options));
            this.logger = logger;
            this.localizer = localizer;
            this.tenantId = this.options.Value.TenantId;
        }

        /// <summary>
        ///  Processes an incoming activity in middleware.
        /// </summary>
        /// <param name="turnContext">The context object for this turn.</param>
        /// <param name="next">The delegate to call to continue the bot middleware pipeline.</param>
        /// <param name="cancellationToken"> A cancellation token that can be used by other objects or threads to receive notice of cancellation.</param>
        /// <returns><see cref="Task"/> A task that represents the work queued to execute.</returns>
        /// <remarks>
        /// Middleware calls the next delegate to pass control to the next middleware in
        /// the pipeline. If middleware doesn’t call the next delegate, the adapter does
        /// not call any of the subsequent middleware’s request handlers or the bot’s receive
        /// handler, and the pipeline short circuits.
        /// The turnContext provides information about the incoming activity, and other data
        /// needed to process the activity.
        /// </remarks>
        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = default)
        {
            turnContext = turnContext ?? throw new ArgumentNullException(nameof(turnContext));
            next = next ?? throw new ArgumentNullException(nameof(next));

            // An explicit check is required for activity type: 'Event' for requests coming from client application extensions.
            // Refer https://github.com/Microsoft/botframework-sdk/blob/master/specs/botframework-activity/botframework-activity.md#event-activity
            // To read more about Bot Framework -- Activity
            if (turnContext.Activity.Type != ActivityTypes.Event && !this.IsActivityFromExpectedTenant(turnContext))
            {
                this.logger.LogWarning($"Unexpected tenant id {turnContext.Activity?.Conversation?.TenantId}");

                if (turnContext.Activity.Type == ActivityTypes.Message)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text(this.localizer.GetString("InvalidTenantText")));
                }

                return;
            }
            else
            {
                await next(cancellationToken);
            }
        }

        /// <summary>
        /// Verify if the tenant Id in the message is the same tenant Id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>True if context is from expected tenant else false.</returns>
        private bool IsActivityFromExpectedTenant(ITurnContext turnContext)
        {
            return turnContext.Activity?.Conversation?.TenantId == this.tenantId;
        }
    }
}
