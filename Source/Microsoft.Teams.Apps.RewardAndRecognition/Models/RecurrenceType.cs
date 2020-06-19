// <copyright file="RecurrenceType.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.RewardAndRecognition.Models
{
    /// <summary>
    /// Enum to specify the award cycle range of recurrence state.
    /// </summary>
    public enum RecurrenceType
    {
        /// <summary>
        /// Single occurrence award cycle.
        /// </summary>
        SingleOccurrence = 0,

        /// <summary>
        /// To repeat award cycle until end of time.
        /// </summary>
        RepeatIndefinitely = 1,

        /// <summary>
        /// To reward cycle until recurrence end date.
        /// </summary>
        RepeatUntilEndDate = 2,

        /// <summary>
        /// To repeat reward cycle occurring N number of times.
        /// </summary>
        RepeatUntilOccurrenceCount = 3,
    }
}
