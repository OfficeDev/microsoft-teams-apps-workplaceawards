// <copyright file="NominationReminderNotificationService.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.BackgroundService
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Hosting;
    using Microsoft.Extensions.Logging;

    /// <summary>
    /// This class inherits BackgroundService base class for implementing long running IHostedServices.
    /// The class implements the background task for sending nomination reminder notifications.
    /// </summary>
    public class NominationReminderNotificationService : BackgroundService
    {
        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<NominationReminderNotificationService> logger;

        /// <summary>
        /// Instance of background service helper.
        /// </summary>
        private readonly INominationReminderBackgroundServiceHelper backgroundServiceHelper;

        /// <summary>
        /// Initializes a new instance of the <see cref="NominationReminderNotificationService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to notification.
        /// </summary>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="backgroundServiceHelper">Helper to send notification.</param>
        public NominationReminderNotificationService(ILogger<NominationReminderNotificationService> logger, INominationReminderBackgroundServiceHelper backgroundServiceHelper)
        {
            this.logger = logger;
            this.backgroundServiceHelper = backgroundServiceHelper;
        }

        /// <summary>
        /// Method to start the background task when application starts.
        /// </summary>
        /// <param name="stoppingToken">A cancellation token that can be used to receive notice of cancellation.</param>
        /// <returns>A task that sends notification in channel for notification.</returns>
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    this.logger.LogInformation("Check and send nomination reminder notification execution started...");
                    await this.backgroundServiceHelper.SendNominationReminderNotificationAsync();

                    this.logger.LogInformation("Check and send nomination reminder notification executed successfully.");
                }
#pragma warning disable CA1031 // Catching general exceptions that might arise during execution to avoid blocking next run.
                catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions that might arise during execution to avoid blocking next run.
                {
                    this.logger.LogError(ex, "Error occurred while running reminder notification service.");
                }
                finally
                {
                    await Task.Delay(TimeSpan.FromHours(12), stoppingToken);
                }
            }

            this.logger.LogInformation("Nomination reminder background service execution ended.");
        }
    }
}
