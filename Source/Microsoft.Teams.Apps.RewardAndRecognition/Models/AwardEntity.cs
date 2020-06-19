// <copyright file="AwardEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using System;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Serialization;

    /// <summary>
    /// Class contains award details created for a team.
    /// </summary>
    public class AwardEntity : TableEntity
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
        /// Gets or sets award id.
        /// </summary>
        [JsonProperty("AwardId")]
        public string AwardId
        {
            get { return this.RowKey; }
            set { this.RowKey = value; }
        }

        /// <summary>
        /// Gets or sets award name.
        /// </summary>
        [JsonProperty("AwardName")]
        public string AwardName { get; set; }

        /// <summary>
        /// Gets or sets award description.
        /// </summary>
        public string AwardDescription { get; set; }

        /// <summary>
        /// Gets or sets award image URL.
        /// </summary>
        public string AwardLink { get; set; }

        /// <summary>
        /// Gets or sets the Azure Active Directory Id of the user created the award.
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets the created by admin principal name.
        /// </summary>
        public string CreatedByAdminUserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of admin.
        /// </summary>
        public string CreatedByAdminObjectId { get; set; }

        /// <summary>
        /// Gets or sets the date time when the award was created.
        /// </summary>
        public DateTime CreatedOn { get; set; }

        /// <summary>
        /// Gets or sets the user modified the award.
        /// </summary>
        public string ModifiedBy { get; set; }

        /// <summary>
        /// Gets or sets the modified by admin principal name.
        /// </summary>
        public string ModifiedByAdminUserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets AAD object id of modified by admin.
        /// </summary>
        public string ModifiedByAdminObjectId { get; set; }

        /// <summary>
        /// Gets or sets the date time when the award was modified.
        /// </summary>
        public DateTime? ModifiedOn { get; set; }
    }
}
