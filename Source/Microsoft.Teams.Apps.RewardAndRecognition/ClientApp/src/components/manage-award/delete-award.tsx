// <copyright file="delete-award.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Text, Flex, Button, Image, Table } from "@fluentui/react-northstar";
import { createBrowserHistory } from "history";
import * as microsoftTeams from "@microsoft/teams-js";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { deleteSelectedAwards } from "../../api/awards-api";
import { withTranslation, WithTranslation } from "react-i18next";
import { IAwardData } from "../../models/award";
import { getBaseUrl } from "../../helpers/utility";

const browserHistory = createBrowserHistory({ basename: "" });

interface IAwardProps extends WithTranslation {
    awardsData: IAwardData[],
    selectedAwards: string[],
    teamId: string,
    onBackButtonClick: () => void,
    onSuccess: (operation: string) => void
}

interface IAwardState {
    error: string,
    isSubmitLoading: boolean;
}

class DeleteAward extends React.Component<IAwardProps, IAwardState> {
    telemetry?: any = null;
    appInsights: any;
    userObjectId?: string = "";

    constructor(props: any) {
        super(props);

        this.state = {
            error: "",
            isSubmitLoading: false,
        }

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.appInsights = {};
    }

    /**
   * Used to initialize Microsoft Teams sdk
   */
    async componentDidMount() {

        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;

            // Initialize application insights for logging events and errors.
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
        });
    }

    /**
   * Handles delete award event.
   */
    onDeleteButtonClick = async () => {
        this.appInsights.trackTrace({ message: `Delete award - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        this.setState({ isSubmitLoading: true });
        let awardIds = this.props.selectedAwards.join(',');
        let deletionResult = await deleteSelectedAwards(awardIds, this.props.teamId!);

        if (deletionResult.data) {
            this.appInsights.trackTrace({ message: `'Delete award' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            this.appInsights.trackEvent({ name: `Delete award` }, { User: this.userObjectId, Team: this.props.teamId! });
            this.props.onSuccess("delete");
        }
        else {
            this.appInsights.trackTrace({ message: `'Delete award' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            this.setState({ error: deletionResult.statusText, isSubmitLoading: true })
        }
    }

    /**
   * Returns awards that are to be deleted.
   */
    deleteContent = () => {
        let awards = this.props.awardsData.filter((award) => {
            return this.props.selectedAwards.includes(award.AwardId);
        });

        let awardsTableRows = awards.map((value: any, index) => (
            {
                key: index,
                style: {},
                items:
                    [
                        {
                            content: <Image alt="NA" fluid src={(value.awardLink === null || value.awardLink === "") ? getBaseUrl() + "/content/DefaultAwardImage.png" : value.awardLink} />, key: index + "1", truncateContent: true, className: "award-image-icon table-image-cell"
                        },
                        {
                            content: <Flex column>
                                <Text content={value.AwardName} weight="semibold" title={value.AwardName} />
                                <Text content={value.awardDescription} title={value.awardDescription} />
                            </Flex>, key: index + "2", truncateContent: true
                        }
                    ]
            }));

        return awardsTableRows;
    }

    render() {
        const { t } = this.props;

        return (
            <>
                <div className="tab-container">
                    <Flex hAlign="center" className="margin-medium-top">
                        <Text content={this.state.error} className="field-error-message" error size="medium" />
                    </Flex>
                    <Text weight="semibold" className="nominee-margin" content={t('awardDeleteConfirmationMessageText')} />
                    <Table className="title" rows={this.deleteContent()} />
                </div>
                <div className="tab-footer">
                    <div>
                        <Flex space="between">
                            <Button icon="icon-chevron-start"
                                content={t('backButtonText')} text
                                onClick={this.props.onBackButtonClick} />
                            <Flex gap="gap.small">
                                <Button content={t('deleteButtonText')} primary
                                    loading={this.state.isSubmitLoading}
                                    onClick={() => { this.onDeleteButtonClick() }}
                                />
                            </Flex>
                        </Flex>
                    </div>
                </div>
            </>
        );
    }
}

export default withTranslation()(DeleteAward)