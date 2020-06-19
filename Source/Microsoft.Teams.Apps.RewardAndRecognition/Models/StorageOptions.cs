// <copyright file="StorageOptions.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    /// <summary>
    /// Provides application setting related to Azure table storage.
    /// </summary>
    public class StorageOptions
    {
        /// <summary>
        /// Gets or sets Azure table storage connection string.
        /// </summary>
        public string ConnectionString { get; set; }
    }
}
