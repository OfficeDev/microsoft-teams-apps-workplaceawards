/*
    <copyright file="preview-nominated-award.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import { Text, Flex, Image, Header, Button } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { NominationAwardPreview } from "../../models/nomination-award-preview";
import { createBrowserHistory } from "history";
import "../../styles/site.css";
import { NominateEntity } from "../../models/nominate-entity";
import Constants from "../../constants/constants";
import { useTranslation } from 'react-i18next';
import { saveNominateDetails, checkDuplicateNomination } from "../../api/nominate-awards-api";
import { useState } from "react";
import { getRewardCycle } from "../../api/reward-cycle-api";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";

interface INominatedAwardProps {
    NominationAwardPreview: NominationAwardPreview,
    onBackButtonClick: () => void,
};

const browserHistory = createBrowserHistory({ basename: "" });

/** Component for previewing award created before sharing in team. */
const PreviewAward = (props: INominatedAwardProps): JSX.Element => {

    const { t } = useTranslation();
    const [isSubmitLoading, setSubmitLoading] = useState<boolean | false | undefined>(false);
    const [errorMessage, setErrorMessage] = useState<string | null | undefined>(null);
    const telemetry = props.NominationAwardPreview.telemetry;

    // Initialize application insights for logging events and errors.
    let appInsights = getApplicationInsightsInstance(telemetry, browserHistory);

    /**
     * Handle save nominated detail event.
    */
    const saveNominatedDetails = async () => {
        setSubmitLoading(true);

        let cycleId;
        let rewardCycleResponse = await getRewardCycle(props.NominationAwardPreview.TeamId!, true)
        if (rewardCycleResponse.data) {
            cycleId = rewardCycleResponse.data.cycleId;
            appInsights.trackTrace({ message: `'getRewardCycle' - Request success`, properties: { User: props.NominationAwardPreview.NominatedByObjectId }, severityLevel: SeverityLevel.Information });
        }

        let isAlreadyNominated = await checkDuplicateNomination(props.NominationAwardPreview.TeamId!, props.NominationAwardPreview.NomineeObjectIds.join(", "), cycleId, props.NominationAwardPreview.AwardId!, props.NominationAwardPreview.NominatedByObjectId!);
        if (isAlreadyNominated.data) {
            appInsights.trackTrace({ message: `'getNominatedAwarddetails' - Request success`, properties: { User: props.NominationAwardPreview.NominatedByUserPrincipalName }, severityLevel: SeverityLevel.Information });
            setErrorMessage(t('alreadyNominatedMessage'));
            setSubmitLoading(false);

            return;
        }

        let nominateEntity: NominateEntity = {
            AwardId: props.NominationAwardPreview.AwardId,
            RewardCycleId: cycleId,
            AwardName: props.NominationAwardPreview.AwardName,
            AwardImageLink: props.NominationAwardPreview.ImageUrl,
            ReasonForNomination: props.NominationAwardPreview.ReasonForNomination,
            TeamId: props.NominationAwardPreview.TeamId,
            NominatedOn: new Date(),
            NomineeNames: props.NominationAwardPreview.Nominees.join(", "),
            NomineeUserPrincipalNames: props.NominationAwardPreview.NomineeUserPrincipalNames.join(", "),
            NomineeObjectIds: props.NominationAwardPreview.NomineeObjectIds.join(", "),
            NominatedByName: props.NominationAwardPreview.NominatedByName,
            NominatedByUserPrincipalName: props.NominationAwardPreview.NominatedByUserPrincipalName,
            NominatedByObjectId: props.NominationAwardPreview.NominatedByObjectId,
            GroupName: JSON.stringify(props.NominationAwardPreview.Nominees),
        };

        appInsights.trackTrace({ message: `'saveNominatedDetails' - Request initiated`, severityLevel: SeverityLevel.Information, properties: { UserEmail: props.NominationAwardPreview.NominatedByUserPrincipalName } });
        const saveNominateDetailsResponse = await saveNominateDetails(nominateEntity);
        if (saveNominateDetailsResponse.data) {
            appInsights.trackTrace({ message: `'saveNominatedDetails' - Request success`, severityLevel: SeverityLevel.Information, properties: { UserEmail: props.NominationAwardPreview.NominatedByUserPrincipalName } });
            appInsights.trackEvent({ name: `Nominate award` }, { User: props.NominationAwardPreview.NominatedByUserPrincipalName, Team: props.NominationAwardPreview.TeamId });
            let toBot = {
                Command: Constants.SaveNominationCommand,
                NominatedByName: props.NominationAwardPreview.NominatedByName,
                NomineeNames: props.NominationAwardPreview.Nominees.join(", "),
                NomineeUserPrincipalNames: props.NominationAwardPreview.NomineeUserPrincipalNames.join(", "),
                NomineeObjectIds: props.NominationAwardPreview.NomineeObjectIds.join(", "),
                AwardId: props.NominationAwardPreview.AwardId,
                AwardName: props.NominationAwardPreview.AwardName,
                AwardLink: props.NominationAwardPreview.ImageUrl,
                ReasonForNomination: props.NominationAwardPreview.ReasonForNomination,
                TeamId: props.NominationAwardPreview.TeamId,
                RewardCycleId: cycleId,
                GroupName: JSON.stringify(props.NominationAwardPreview.Nominees),
            };

            microsoftTeams.tasks.submitTask(toBot);
        }
        else {
            setErrorMessage(t('errorMessage'));
            setSubmitLoading(false);
            appInsights.trackTrace({ message: `'saveNominatedDetails' - Request failed`, severityLevel: SeverityLevel.Error, properties: { UserEmail: props.NominationAwardPreview.NominatedByUserPrincipalName, Code: saveNominateDetailsResponse.status } });
        }
    }

    /**
    *  Returns the nominated award preview to parent.
    * */
    return (
        <>
            <div className="preview-container">
                <Flex hAlign="center" className="header-nomination">
                    <Text
                        content={t('previewAwardHeader')}
                    />
                </Flex>
                <div className="div-shadow">
                    <div className="div-shadow-nomination">
                        <Flex gap="gap.medium" padding="padding.medium" hAlign="start" vAlign="start">
                            <Flex.Item align="start" size="size.small" grow push>
                                <Flex space="between">
                                    <Text weight="bold" className="word-break award-cycle-header" content={props.NominationAwardPreview.AwardName} />
                                </Flex>
                            </Flex.Item>
                            <div className="image-size-alignment">
                                <Flex.Item align="start" size="size.small">
                                    <Image className="preview-image-icon" fluid src={props.NominationAwardPreview.ImageUrl} />
                                </Flex.Item>
                            </div>
                        </Flex>
                        <Flex gap="gap.medium" padding="padding.medium" hAlign="start" vAlign="center">
                            <Flex.Item align="start" size="size.small" grow>
                                <Flex column gap="gap.small" vAlign="stretch">
                                    <Flex space="between">
                                        <Text className="nominee-margin" weight="bold" content={props.NominationAwardPreview.Nominees.join(", ")} />
                                    </Flex>
                                </Flex>
                            </Flex.Item>
                        </Flex>
                        <Flex gap="gap.medium" padding="padding.medium" hAlign="start" vAlign="center">
                            <Flex.Item align="start" size="size.small" grow>
                                <Flex column gap="gap.small" vAlign="stretch">
                                    <Flex space="between">
                                        <Text content={t('nominatedByText') + props.NominationAwardPreview.NominatedByName} />
                                    </Flex>
                                </Flex>
                            </Flex.Item>
                        </Flex>
                        <Flex gap="gap.medium" padding="padding.medium" hAlign="start" vAlign="center">
                            <Flex.Item align="start" size="size.small" grow>
                                <Flex column gap="gap.small" vAlign="stretch">
                                    <Flex space="between">
                                        <Text className="word-break" content={props.NominationAwardPreview.ReasonForNomination} />
                                    </Flex>
                                </Flex>
                            </Flex.Item>
                        </Flex>
                    </div>
                </div>
            </div>
            <Flex gap="gap.small">
                {errorMessage !== null && <Text className="small-margin-left" content={errorMessage} error />}
            </Flex>
            <div className="tab-footer">
                <div>
                    <Flex space="between">
                        <Button icon="icon-chevron-start"
                            content={t('backButtonText')} text
                            onClick={props.onBackButtonClick}
                        />
                        <Flex gap="gap.small">
                            <Button content={t('nominateButton')} primary className="preview-button"
                                loading={isSubmitLoading}
                                disabled={isSubmitLoading}
                                onClick={() => { saveNominatedDetails() }}
                            />
                        </Flex>
                    </Flex>
                </div>
            </div>
        </>
    );
};

export default PreviewAward;