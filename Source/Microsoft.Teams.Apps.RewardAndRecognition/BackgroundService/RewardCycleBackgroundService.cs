// <copyright file="RewardCycleBackgroundService.cs" company="Microsoft">
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
    /// The class implements the background task to update reward cycles in storage.
    /// </summary>
    public class RewardCycleBackgroundService : BackgroundService
    {
        /// <summary>
        /// Instance of background service helper.
        /// </summary>
        private readonly IRewardCycleBackgroundServiceHelper backgroundServiceHelper;

        /// <summary>
        /// Instance to send logs to the Application Insights service.
        /// </summary>
        private readonly ILogger<RewardCycleBackgroundService> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="RewardCycleBackgroundService"/> class.
        /// BackgroundService class that inherits IHostedService and implements the methods related to award cycle.
        /// </summary>
        /// <param name="backgroundServiceHelper">Helper to update reward cycle.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        public RewardCycleBackgroundService(IRewardCycleBackgroundServiceHelper backgroundServiceHelper, ILogger<RewardCycleBackgroundService> logger)
        {
            this.backgroundServiceHelper = backgroundServiceHelper;
            this.logger = logger;
        }

        /// <summary>
        /// Method to start the background task when application starts.
        /// </summary>
        /// <param name="stoppingToken">A cancellation token that can be used to receive notice of cancellation.</param>
        /// <returns>A task that sets reward cycle.</returns>
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    this.logger.LogInformation("Check and update reward cycle execution check started...");

                    await this.backgroundServiceHelper.UpdateCycleStatusAsync();

                    this.logger.LogInformation("Check and update reward cycle execution completed");
                }
#pragma warning disable CA1031 // Catching general exceptions that might arise while updating reward cycle state to avoid blocking next execution.
                catch (Exception ex)
#pragma warning restore CA1031 // Catching general exceptions that might arise while updating reward cycle state to avoid blocking next execution.
                {
                    this.logger.LogError(ex, "Error while updating reward cycle from background service.");
                }
                finally
                {
                    await Task.Delay(TimeSpan.FromHours(4), stoppingToken);
                }
            }

            this.logger.LogInformation("Check and update reward cycle background service execution ended.");
        }
    }
}
