// <copyright file="JwtClaims.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    /// <summary>
    /// Claims which are added in JWT token.
    /// </summary>
    public class JwtClaims
    {
        /// <summary>
        /// Gets or sets activity Id.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets service URL of bot.
        /// </summary>
        public string ServiceUrl { get; set; }

        /// <summary>
        /// Gets or sets valid upn.
        /// </summary>
        public string Upn { get; set; }
    }
}
