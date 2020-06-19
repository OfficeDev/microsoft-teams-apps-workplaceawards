// <copyright file="PolicyNames.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Authentication.AuthenticationPolicy
{
    /// <summary>
    /// This class list the policy name of custom authorizations implemented in project.
    /// </summary>
    public static class PolicyNames
    {
        /// <summary>
        /// The name of the authorization policy, MustBeTeamCaptainUserPolicy.
        /// A team champion has permission to set up awards, nomination and publish results.
        /// </summary>
        public const string MustBeTeamCaptainUserPolicy = "MustBeTeamCaptainUserPolicy";

        /// <summary>
        /// The name of the authorization policy, MustBeTeamMemberUserPolicy.
        /// Indicates that user is a part of team and has permission to nominate and endorse team members.
        /// </summary>
        public const string MustBeTeamMemberUserPolicy = "MustBeTeamMemberUserPolicy";
    }
}
