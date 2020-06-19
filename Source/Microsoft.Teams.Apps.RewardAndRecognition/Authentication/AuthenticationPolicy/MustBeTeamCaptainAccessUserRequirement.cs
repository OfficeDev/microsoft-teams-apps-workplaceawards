// <copyright file="MustBeTeamCaptainAccessUserRequirement.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Authentication.AuthenticationPolicy
{
    using Microsoft.AspNetCore.Authorization;

    /// <summary>
    /// This authorization class implements the marker interface
    /// <see cref="IAuthorizationRequirement"/> to check if user meets team captain specific requirements
    /// for accesing resources.
    /// </summary>
    public class MustBeTeamCaptainAccessUserRequirement : IAuthorizationRequirement
    {
    }
}
