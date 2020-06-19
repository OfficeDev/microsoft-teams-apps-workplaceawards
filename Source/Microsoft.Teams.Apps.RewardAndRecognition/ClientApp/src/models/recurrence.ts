/*
    <copyright file="recurrence.ts" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

/// 0 = Non recursive award cycle.
/// 1 = To repeat award cycle until end of time.
/// 2 = To reward cycle until recurrence end date.
/// 3 = To repeat reward cycle occurring N number of times.
export enum Recurrence {
    SingleOccurrence = 0,
    RepeatIndefinitely = 1,
    RepeatUntilEndDate = 2,
    RepeatUntilOccurrenceCount = 3
}