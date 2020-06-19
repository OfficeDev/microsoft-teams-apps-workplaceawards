// <copyright file="RewardAndRecognitionActivityHandlerOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition
{
    /// <summary>
    /// The RewardAndRecognitionActivityHandlerOptions are the options for the <see cref="RewardAndRecognitionActivityHandler" /> bot.
    /// </summary>
    public sealed class RewardAndRecognitionActivityHandlerOptions
    {
        /// <summary>
        /// Gets or sets unique id of Tenant.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets application base Uri.
        /// </summary>
        public string AppBaseUri { get; set; }
    }
}