// <copyright file="BaseRewardAndRecognitionController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Controllers
{
    using System.Linq;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.RewardAndRecognition.Models;

    /// <summary>
    /// This ASP controller act as a base controller.
    /// It is created to handle incoming request and provides implementation to get user claims.
    /// Inherits <see cref="ControllerBase"/> is a base class for an MVC controller without view support.
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class BaseRewardAndRecognitionController : ControllerBase
    {
        /// <summary>
        /// Get claims of user.
        /// </summary>
        /// <returns>User claims.</returns>
        protected JwtClaims GetUserClaims()
        {
            var claims = this.User.Claims;
            var jwtClaims = new JwtClaims
            {
                FromId = claims.Where(claim => claim.Type == "http://schemas.microsoft.com/identity/claims/objectidentifier").Select(claim => claim.Value).First(),
                Upn = claims.Where(claim => claim.Type == "http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn").Select(claim => claim.Value).First(),
            };

            return jwtClaims;
        }
    }
}