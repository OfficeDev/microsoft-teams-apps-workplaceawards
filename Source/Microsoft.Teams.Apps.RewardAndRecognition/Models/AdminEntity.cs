// <copyright file="AdminEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using System;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Class contains details of award admin. Awards admin can define awards, set awards cycle, and share results.
    /// </summary>
    public class AdminEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets team id.
        /// </summary>
        public string TeamId
        {
            get { return this.PartitionKey; }
            set { this.PartitionKey = value; }
        }

        /// <summary>
        /// Gets or sets unique identifier of row.
        /// </summary>
        public string RowUniqueId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets name of user who configured the admin for the team.
        /// </summary>
        [JsonProperty("CreatedByUserPrincipalName")]
        public string CreatedByUserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of user who configured the admin.
        /// </summary>
        [JsonProperty("CreatedByObjectId")]
        public string CreatedByObjectId { get; set; }

        /// <summary>
        /// Gets or sets date on when the admin is configured.
        /// </summary>
        [JsonProperty("CreatedOn")]
        public DateTime? CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets name of admin name for the team.
        /// </summary>
        [JsonProperty("AdminName")]
        public string AdminName { get; set; }

        /// <summary>
        /// Gets or sets admin user principal name.
        /// </summary>
        [JsonProperty("AdminUserPrincipalName")]
        public string AdminUserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of admin.
        /// </summary>
        [JsonProperty("AdminObjectId")]
        public string AdminObjectId { get; set; }

        /// <summary>
        /// Gets or sets Note that was given to the team.
        /// </summary>
        [JsonProperty("NoteForTeam")]
        public string NoteForTeam { get; set; }
    }
}
