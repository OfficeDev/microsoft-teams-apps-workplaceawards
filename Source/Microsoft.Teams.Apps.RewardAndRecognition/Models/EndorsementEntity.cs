// <copyright file="EndorsementEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    using System;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Class contains nomination endorsement details of team members.
    /// A member can support team members on nominated awards.
    /// </summary>
    public class EndorsementEntity : TableEntity
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
        /// Gets or sets endorsed award name.
        /// </summary>
        public string EndorsedForAward { get; set; }

        /// <summary>
        /// Gets or sets endorsed award id.
        /// </summary>
        public string EndorsedForAwardId { get; set; }

        /// <summary>
        /// Gets or sets award cycle.
        /// </summary>
        public string AwardCycle { get; set; }

        /// <summary>
        /// Gets or sets endorsee user principal name.
        /// </summary>
        public string EndorseeUserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets the Azure Active Directory Id of the endorsee.
        /// </summary>
        public string EndorseeObjectId { get; set; }

        /// <summary>
        /// Gets or sets the endorsed by principal name.
        /// </summary>
        public string EndorsedByUserPrincipalName { get; set; }

        /// <summary>
        /// Gets or sets Azure Active Directory object id of endorsed by user.
        /// </summary>
        public string EndorsedByObjectId { get; set; }

        /// <summary>
        /// Gets or sets the date time when the award was endorsed.
        /// </summary>
        public DateTime EndorsedOn { get; set; }
    }
}
