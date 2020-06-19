// <copyright file="reward-cycle.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import { Button, Checkbox, Flex, Input, RadioGroup, Text, Loader, Icon } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { Fabric, Customizer } from 'office-ui-fabric-react/lib';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import * as React from "react";
import { useEffect, useState } from "react";
import { useTranslation } from "react-i18next";
import { getRewardCycle, setRewardCycle } from "../../api/reward-cycle-api";
import { createBrowserHistory } from "history";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import { Recurrence } from "../../models/recurrence";
import { RewardCycleDetail } from "../../models/reward-cycle";
import Constants from "../../constants/constants";
import { DarkCustomizations } from "../../helpers/theme/DarkCustomizations";
import { DefaultCustomizations } from "../../helpers/theme/DefaultCustomizations";
let moment = require('moment');

initializeIcons();

const browserHistory = createBrowserHistory({ basename: "" });

interface IRewardCycleState {
    selectedValue: number,
    noOfOccurence: string | undefined,
    isReccurringChecked: boolean,
    error: string
}

interface ICycle {
    cycleId: string | undefined;
    rewardCycleStartDate: Date | null | undefined;
    rewardCycleEndDate: Date | null | undefined;
    numberOfOccurrences: number | undefined;
    teamId: string | undefined;
    isRecurring: number | undefined;
    recurrence: number | undefined;
    rangeOfOccurrenceEndDate: Date | null | undefined;
    cycleStatus: number | undefined;
    CreatedByUserPrincipalName: string | undefined;
    createdByObjectId: string | undefined;
    createdOn: Date | null | undefined;
    resultPublished: number | undefined;
    isCycleActive: boolean | false;
}

interface IProps {
    teamId: string,
}

const controlClass = mergeStyleSets({
    control: {
        margin: '0 0 15px 0',
        maxWidth: '300px'
    }
});

const RewardCycle: React.FC<IProps> = props => {
    let search = window.location.search;
    let params = new URLSearchParams(search);
    let telemetry = params.get("telemetry");
    let appInsights = getApplicationInsightsInstance(telemetry, browserHistory);
    let userObjectId: string | undefined;
    let userEmail: string | undefined;
    const { t } = useTranslation();
    const [startDate, setStartDate] = useState<Date | null | undefined>(null);
    const [endDate, setEndDate] = useState<Date | null | undefined>(null);
    const [minEndDate, setMinEndDate] = useState<Date>(new Date(moment().add(Constants.minimumCycleDays, 'd').format()));
    const [calendarDate, setCalendarDate] = useState<Date | null | undefined>(null);
    const [summary, setSummary] = useState<string>('');
    const [rewardCycleState, setRewardCycleState] =
        useState<IRewardCycleState>({ selectedValue: Recurrence.RepeatIndefinitely, noOfOccurence: '', isReccurringChecked: false, error: '' });
    const [cycleState, setCycleState] =
        useState<ICycle>({
            cycleId: undefined,
            createdByObjectId: '',
            CreatedByUserPrincipalName: '',
            createdOn: undefined,
            isRecurring: undefined,
            numberOfOccurrences: undefined,
            recurrence: undefined,
            rangeOfOccurrenceEndDate: undefined,
            resultPublished: 0,
            rewardCycleEndDate: undefined,
            rewardCycleStartDate: undefined,
            cycleStatus: undefined,
            teamId: '',
            isCycleActive: false,
        });
    const [loader, setLoader] = useState(true);
    const [submitLoading, setSubmitLoading] = useState(false);
    const [datePickerTheme, setDatePickerTheme] = useState(DefaultCustomizations);
    useEffect(() => {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            userObjectId = context.userObjectId;
            userEmail = context.upn;
            let themeContext = context.theme || "";
            if (themeContext === Constants.dark) { setDatePickerTheme(DarkCustomizations) }
            else if (themeContext === Constants.contrast) { setDatePickerTheme(DarkCustomizations) }
            else { setDatePickerTheme(DefaultCustomizations) }
        });
        const fetchData = async () => {
            appInsights.trackTrace({ message: `'getRewardCycle' - Initiated request`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
            let response = await getRewardCycle(props.teamId, true)
            if (response.data) {
                let rewardcycle = response.data;
                appInsights.trackTrace({ message: `'getRewardCycle' - Request success`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
                setCycleState({
                    cycleId: rewardcycle.cycleId,
                    createdByObjectId: rewardcycle.createdByObjectId,
                    CreatedByUserPrincipalName: rewardcycle.CreatedByUserPrincipalName,
                    createdOn: rewardcycle.createdOn,
                    isRecurring: rewardcycle.recurrence > 0 ? 1 : 0,
                    numberOfOccurrences: rewardcycle.numberOfOccurrences,
                    recurrence: rewardcycle.recurrence,
                    rangeOfOccurrenceEndDate: rewardcycle.rangeOfOccurrenceEndDate,
                    resultPublished: rewardcycle.resultPublished,
                    rewardCycleEndDate: rewardcycle.rewardCycleEndDate,
                    rewardCycleStartDate: rewardcycle.rewardCycleStartDate,
                    cycleStatus: rewardcycle.rewardCycleState,
                    teamId: props.teamId,
                    isCycleActive: true,
                });
                if (rewardcycle.rewardCycleStartDate) {
                    setStartDate(new Date(rewardcycle.rewardCycleStartDate));
                    setMinEndDate(new Date(moment(rewardcycle.rewardCycleStartDate).add(Constants.minimumCycleDays, 'd').format()));
                }
                if (rewardcycle.rewardCycleEndDate) { setEndDate(new Date(rewardcycle.rewardCycleEndDate)); }
                if (rewardcycle.rangeOfOccurrenceEndDate) {
                    setCalendarDate(new Date(rewardcycle.rangeOfOccurrenceEndDate));
                }
                setRewardCycleState({
                    isReccurringChecked: rewardcycle.recurrence > 0 ? true : false,
                    noOfOccurence: rewardcycle.numberOfOccurrences !== 0 ? rewardcycle.numberOfOccurrences : undefined,
                    selectedValue: rewardcycle.recurrence!,
                    error: ''
                });
                generateCycleSummary(
                    rewardcycle.rewardCycleStartDate,
                    rewardcycle.rewardCycleEndDate,
                    rewardcycle.recurrence > 0 ? true : false,
                    rewardcycle.recurrence,
                    rewardcycle.numberOfOccurrences,
                    rewardcycle.rangeOfOccurrenceEndDate
                );
            }
            else {
                appInsights.trackTrace({ message: `'getRewardCycle' - Request failed`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
            }
            setLoader(false);
        };
        fetchData();
    }, [datePickerTheme]);

    /**
     * Handle change event for cycle start date picker.
     * @param date | cycle start date.
     */
    const onSelectStartDate = (date: Date | null | undefined): void => {
        setStartDate(date);
        setMinEndDate(new Date(moment(date).add(Constants.minimumCycleDays, 'd').format()));
        generateCycleSummary(date, endDate, rewardCycleState.isReccurringChecked, rewardCycleState.selectedValue, parseInt(rewardCycleState.noOfOccurence!), calendarDate);
    };

    /**
     * Handle change event for cycle end date picker.
     * @param date | cycle end date.
     */
    const onSelectEndDate = (date: Date | null | undefined): void => {
        setEndDate(date);
        generateCycleSummary(startDate, date, rewardCycleState.isReccurringChecked, rewardCycleState.selectedValue, parseInt(rewardCycleState.noOfOccurence!), calendarDate);
    };
    /**
     * Handle change event for end by date picker.
     * @param date | end by date.
     */
    const onSelectCalendarDate = (date: Date | null | undefined): void => {
        setCalendarDate(date);
        generateCycleSummary(startDate, endDate, rewardCycleState.isReccurringChecked, rewardCycleState.selectedValue, parseInt(rewardCycleState.noOfOccurence!), date);
    };

    /**
     * Handling input change event.
     * @param event 
     */
    const handleInputChange = (event: any): void => {
        let p = event.target;
        setRewardCycleState({ ...rewardCycleState, [p.name]: p.value })
        generateCycleSummary(startDate, endDate, rewardCycleState.isReccurringChecked, rewardCycleState.selectedValue, parseInt(p.value), calendarDate);
    }

    /**
     * Handling check box change event.
     * @param isChecked | boolean value.
     */
    const handleCheckBoxChange = (isChecked: boolean): void => {
        setRewardCycleState({ isReccurringChecked: !isChecked, selectedValue: Recurrence.RepeatIndefinitely, noOfOccurence: '', error: rewardCycleState.error })
        setCalendarDate(null);
        generateCycleSummary(startDate, endDate, !isChecked, rewardCycleState.selectedValue, parseInt(rewardCycleState.noOfOccurence!), calendarDate);
    }

    /**
     * This method is used to handle done button click by setting the reward cycle and sending notification card to the team.
     * @param start | cycle start date.
     * @param end | cycle end date.
     */
    const onSetCycle = async (start: Date | null | undefined, end: Date | null | undefined): Promise<void> => {
        if (!(start && end)) {
            setRewardCycleState({ ...rewardCycleState, error: t('requiredDatesError') });
            return;
        }
        if (moment(end).diff(moment(start), 'd') < Constants.minimumCycleDays) {
            setRewardCycleState({ ...rewardCycleState, error: t('minimumRewardCycleValidationMessage', { day: Constants.minimumCycleDays }) });
            return;
        }
        if (rewardCycleState.selectedValue === Recurrence.RepeatUntilOccurrenceCount) {
            if (parseInt(rewardCycleState.noOfOccurence!) <= 0) {
                setRewardCycleState({ ...rewardCycleState, error: t('noOfOccurrenceError') });
                return;
            }
            if (!rewardCycleState.noOfOccurence || rewardCycleState.noOfOccurence === '') {
                setRewardCycleState({ ...rewardCycleState, error: t('noOfOccurrenceError') });
                return;
            }
        }
        let startCycle = moment(start)
            .set('hour', moment().hour())
            .set('minute', moment().minute())
            .set('second', moment().second());
        let endCycle = moment(end)
            .set('hour', moment().hour())
            .set('minute', moment().minute())
            .set('second', moment().second());
        let endByDate = undefined;
        if (rewardCycleState.selectedValue === Recurrence.RepeatUntilEndDate) {
            if (!calendarDate || calendarDate === null) {
                setRewardCycleState({ ...rewardCycleState, error: t('requiredEndByDate') });
                return;
            }
            endByDate = moment.utc(moment(calendarDate)
                .set('hour', moment().hour())
                .set('minute', moment().minute())
                .set('second', moment().second()));
        }
        setSubmitLoading(true);
        let rewardCycleDetail: RewardCycleDetail = {
            RewardCycleStartDate: moment.utc(startCycle),
            RewardCycleEndDate: moment.utc(endCycle),
            NumberOfOccurrences: rewardCycleState.selectedValue === Recurrence.RepeatUntilOccurrenceCount ? parseInt(rewardCycleState.noOfOccurence!) : 0,
            ResultPublished: cycleState.resultPublished,
            RewardCycleState: start.getDate() <= new Date().getDate() ? 1 : 0,
            CycleId: cycleState.cycleId,
            CreatedByUserPrincipalName: userEmail,
            Recurrence: rewardCycleState.isReccurringChecked ? rewardCycleState.selectedValue : 0,
            RangeOfOccurrenceEndDate: endByDate,
            TeamId: props.teamId,
            CreatedByObjectId: userObjectId,
            CreatedOn: cycleState.createdOn,
            ResultPublishedOn: null,
        };
        appInsights.trackTrace({ message: `'setRewardCycle' - Initiated request`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
        let response = await setRewardCycle(rewardCycleDetail);
        if (response.data) {
            appInsights.trackTrace({ message: `'setRewardCycle' - Request success`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
            appInsights.trackEvent({ name: `Set reward cycle` }, { User: userObjectId, Team: props.teamId! });
            let toBot = {
                Command: Constants.NominateAwardsCommand,
                RewardCycleStartDate: rewardCycleDetail.RewardCycleStartDate,
                RewardCycleEndDate: rewardCycleDetail.RewardCycleEndDate,
                RewardCycleId: response.data.cycleId,
                TeamId: props.teamId
            };
            let obj = JSON.parse(JSON.stringify(toBot));
            microsoftTeams.tasks.submitTask(obj);
        }
        else {
            appInsights.trackTrace({ message: `'setRewardCycle' - Request failed`, properties: { User: userObjectId }, severityLevel: SeverityLevel.Information });
            setRewardCycleState({ ...rewardCycleState, error: t('errorText') });
            setSubmitLoading(false);
        }
    };

    /**
     * Handle radio group change event.
     * @param e | event
     * @param props | props
     */
    const handleChange = (e: any, props: any) => {
        setRewardCycleState({ noOfOccurence: '', isReccurringChecked: rewardCycleState.isReccurringChecked, selectedValue: props.value, error: '' });
        setCalendarDate(null);
        generateCycleSummary(startDate, endDate, rewardCycleState.isReccurringChecked, props.value, parseInt(rewardCycleState.noOfOccurence!), null);
    }

    /**
     * Radio group items.
     */
    const getItems = () => {
        return [
            {
                key: 'none',
                label: t('noEndDate'),
                value: Recurrence.RepeatIndefinitely,
            },
            {
                key: 'endafter',
                label: (
                    <Flex vAlign="center" gap="gap.small" className="margin-small-top">
                        <Text content={t('endAfter')} />
                        <Input type="number"
                            min={1}
                            name="noOfOccurence"
                            value={rewardCycleState.noOfOccurence!}
                            onChange={handleInputChange}
                            defaultValue={undefined}
                        />
                        <Text content={t('occurrenceText')} />
                    </Flex>
                ),
                value: Recurrence.RepeatUntilOccurrenceCount,
            },
            {
                key: 'endby',
                label: (
                    <Flex vAlign="center" gap="gap.small" className="margin-small-top" >
                        <Text content={t('endBy')} />
                        <Fabric>
                            <Customizer {...datePickerTheme}>
                                <DatePicker
                                    className={controlClass.control}
                                    allowTextInput={true}
                                    showMonthPickerAsOverlay={true}
                                    minDate={endDate!}
                                    isMonthPickerVisible={true}
                                    value={calendarDate!}
                                    onSelectDate={onSelectCalendarDate}
                                />
                            </Customizer>
                        </Fabric>
                    </Flex>
                ),
                value: Recurrence.RepeatUntilEndDate,
            }
        ];
    }

    const generateCycleSummary = (startDate: Date | null | undefined, endDate: Date | null | undefined, isRecurring: boolean, selectedValue: number | undefined, occurenceNumber: number | undefined, endOnDate: Date | null | undefined) => {
        let days;
        if (startDate && endDate) {
            days = moment(endDate).diff(moment(startDate), 'd');
            if (!isRecurring) {
                setSummary(t('summaryNoRecurrence', { number: days }));
            }
            else {
                if (selectedValue === Recurrence.RepeatIndefinitely) {
                    setSummary(t('summaryNoEnd', { number: days, startCycle: moment.utc(startDate).local().format("ll") }));
                }
                else if (selectedValue === Recurrence.RepeatUntilOccurrenceCount && occurenceNumber) {
                    setSummary(t('summaryEndAfter', { number: days, startCycle: moment.utc(startDate).local().format("ll"), endAfter: occurenceNumber }));
                }
                else if (selectedValue === Recurrence.RepeatUntilEndDate && endOnDate) {
                    setSummary(t('summaryEndOn', { number: days, startCycle: moment.utc(startDate).local().format("ll"), endOn: moment.utc(endOnDate).local().format("ll") }));
                }
                else {
                    setSummary('');
                }
            }
        }
    }

    return (
        <div>
            {loader ?
                <div className="tab-container">
                    <Loader />
                </div>
                :
                <div>
                    <div className="tab-container">
                        {rewardCycleState.error && <Flex hAlign="center"><Text content={rewardCycleState.error} error /></Flex>}
                        {cycleState.isCycleActive && <Flex column className="set-cycle-summary-margin-top">
                            <Text content={summary} /><Text content={t('currentRewardCycle')} />
                        </Flex>}
                        <Flex gap="gap.small" className="header-nomination">
                            <Flex.Item className="margin-large-right">
                                <div>
                                    <Flex gap="gap.small">
                                        <Text content={t('startDate')} /><Text content="*" className="requiredfield" error size="medium" />
                                        <Icon name="info" outline title={t('informationMessageStartDate')} />
                                    </Flex>
                                    <Flex className="margin-small-top">
                                        <Fabric>
                                            <Customizer {...datePickerTheme}>
                                                <DatePicker
                                                    className={controlClass.control}
                                                    allowTextInput={true}
                                                    showMonthPickerAsOverlay={true}
                                                    minDate={new Date()}
                                                    isMonthPickerVisible={true}
                                                    value={startDate!}
                                                    onSelectDate={onSelectStartDate}
                                                />
                                            </Customizer>
                                        </Fabric>
                                    </Flex>
                                </div>
                            </Flex.Item>
                            <Flex.Item className="margin-large-right">
                                <div>
                                    <Flex gap="gap.small">
                                        <Text content={t('endDate')} /><Text content="*" className="requiredfield" error size="medium" />
                                        <Icon name="info" outline title={t('informationMessageEndDate')} />
                                    </Flex>
                                    <Flex className="margin-small-top">
                                        <Fabric >
                                            <Customizer {...datePickerTheme}>
                                                <DatePicker
                                                    className={controlClass.control}
                                                    allowTextInput={true}
                                                    minDate={minEndDate}
                                                    isMonthPickerVisible={true}
                                                    showMonthPickerAsOverlay={true}
                                                    value={endDate!}
                                                    onSelectDate={onSelectEndDate}
                                                />
                                            </Customizer>
                                        </Fabric>
                                    </Flex>
                                </div>
                            </Flex.Item>
                            <Flex.Item >
                                <Flex column gap="gap.small">
                                    <div>
                                        <Text content={t('recurring')} />&nbsp;
                                        <Icon name="info" outline title={t('informationMessageRecurrence')} />
                                    </div>
                                    <Checkbox toggle
                                        checked={rewardCycleState.isReccurringChecked}
                                        onChange={() => handleCheckBoxChange(rewardCycleState.isReccurringChecked)}
                                    />
                                </Flex>
                            </Flex.Item>
                        </Flex>
                        {rewardCycleState.isReccurringChecked && <Flex column gap="gap.small">
                            <div className="set-cycle-margin-top">
                                <Text content={t('rangeOfOccurences')} />
                                <RadioGroup vertical className="margin-medium-top"
                                    defaultCheckedValue={rewardCycleState.selectedValue}
                                    items={getItems()}
                                    onCheckedValueChange={handleChange}
                                />
                            </div>
                        </Flex>}
                        {(startDate && endDate) &&
                            <Flex column className="set-cycle-summary-margin-top">
                            <Text weight="bold" content={t('rewardCycleSummaryTitle')} />
                                <Text content={summary} />
                            </Flex>}
                    </div>
                    <div className="tab-footer">
                        <Flex hAlign="end">
                            <Button primary content={t('doneButtonText')} onClick={() => onSetCycle(startDate, endDate)} loading={submitLoading} />
                        </Flex>
                    </div>
                </div>
            }
        </div>
    );
}

export default RewardCycle;