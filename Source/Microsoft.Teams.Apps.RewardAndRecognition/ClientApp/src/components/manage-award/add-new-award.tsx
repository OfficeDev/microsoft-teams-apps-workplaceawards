// <copyright file="add-new-award.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Text, Flex, Input, TextArea, Button } from "@fluentui/react-northstar";
import { createBrowserHistory } from "history";
import * as microsoftTeams from "@microsoft/teams-js";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { postAward } from "../../api/awards-api";
import { AwardDetails } from "../../models/award";
import { withTranslation, WithTranslation } from "react-i18next";
import { isNullorWhiteSpace, checkUrl } from "../../helpers/utility";

const browserHistory = createBrowserHistory({ basename: "" });

interface IAwardState {
    awardName: string;
    awardDescription: string;
    awardImageLink: string;
    isNameValuePresent: boolean,
    isDescriptionValuePresent: boolean,
    error: string,
    isSubmitLoading: boolean,
    invalidLinkText: string,
}

interface IAwardProps extends WithTranslation {
    awards: Array<any>,
    isNewAllowed: boolean,
    teamId: string,
    onBackButtonClick: () => void,
    onSuccess: (operation: string) => void
}

class AddAward extends React.Component<IAwardProps, IAwardState> {
    telemetry?: any = null;
    theme?: any = null;
    locale?: string | null;
    appInsights: any;
    userObjectId?: string = "";

    constructor(props: any) {
        super(props);

        this.state = {
            awardName: "",
            awardDescription: "",
            awardImageLink: "",
            isNameValuePresent: true,
            isDescriptionValuePresent: true,
            error: "",
            isSubmitLoading: false,
            invalidLinkText: "",
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
            this.theme = context.theme;
            this.locale = context.locale;
            // Initialize application insights for logging events and errors.
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
        });
    }

    /**
   *Checks whether all validation conditions are matched before user submits new response
   */
    checkIfSubmitAllowed = (t: any) => {
        if (isNullorWhiteSpace(this.state.awardName)) {
            this.setState({ isNameValuePresent: false });
            return false;
        }

        if (isNullorWhiteSpace(this.state.awardDescription)) {
            this.setState({ isDescriptionValuePresent: false });
            return false;
        }

        if (this.state.awardName && this.state.awardDescription) {
            let filteredData = this.props.awards.filter((award) => {
                return (award.AwardName.toUpperCase() === this.state.awardName.trim().toUpperCase());
            });

            if (filteredData.length > 0) {
                this.setState({ error: t('duplicateAwardError') })

                return false;
            }
            if (!isNullorWhiteSpace(this.state.awardImageLink)) {

                let result = checkUrl(this.state.awardImageLink);
                if (!result) { this.setState({ invalidLinkText: t('invalidImageLink') }) }

                return result;
            }

            return true;
        }
        else {
            return false;
        }
    }

    /**
     * Handle add award event.
     */
    onAddButtonClick = async (t: any) => {
        if (this.checkIfSubmitAllowed(t)) {
            this.setState({ isSubmitLoading: true });
            let awardDetail: AwardDetails = {
                AwardId: undefined,
                AwardName: this.state.awardName.trim(),
                AwardDescription: this.state.awardDescription.trim(),
                AwardLink: this.state.awardImageLink,
                TeamId: this.props.teamId,
                CreatedBy: undefined,
                CreatedOn: undefined
            };

            this.appInsights.trackTrace({ message: `'addAward' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });            
            let response = await postAward(awardDetail);

            if (response.data) {
                this.appInsights.trackTrace({ message: `'addAward' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
                this.appInsights.trackEvent({ name: `Add award` }, { User: this.userObjectId, Team: this.props.teamId });
                this.setState({ error: '', isSubmitLoading: false });
                this.props.onSuccess("add");
                return;
            }
            else {
                this.appInsights.trackTrace({ message: `'addAward' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
                this.setState({ error: response.statusText, isSubmitLoading: false })
            }
        }
    }

    /**
     * Handle name change event.
     */
    handleInputNameChange = (event: any) => {
        this.setState({ awardName: event.target.value, isNameValuePresent: true, error: "" });
    }

    /**
     * Handle description change event.
     */
    handleInputDescriptionChange = (event: any) => {
        this.setState({ awardDescription: event.target.value, isDescriptionValuePresent: true });
    }

    /**
     * Handle award link change event.
     */
    handleInputImageChange = (event: any) => {        
        this.setState({ awardImageLink: (event.target.value !== "" || event.target.value !== null) ? event.target.value.trim() : null });
    }

    /**
    *Returns text component containing error message for failed name field validation
    *@param {boolean} isValuePresent Indicates whether value is present
    */
    private getRequiredFieldError = (isValuePresent: boolean, t: any) => {
        if (!isValuePresent) {
            return (<Text content={t('fieldRequiredMessage')} className="field-error-message" error size="medium" />);
        }

        return (<></>);
    }

    render() {
        const { t } = this.props;

        return (
            <>
                <div className="tab-container">
                    <div>
                        <Flex hAlign="center">
                            <Text content={this.state.error} className="field-error-message" error size="medium" />
                        </Flex>
                        <Flex gap="gap.small" className="margin-medium-top">
                            <Text content={t('awardName')} size="medium" /><Text content="*" className="requiredfield" error size="medium" />
                            <Flex.Item push>
                                {this.getRequiredFieldError(this.state.isNameValuePresent, t)}
                            </Flex.Item>
                        </Flex>
                        <div className="add-form-input">
                            <Input placeholder={t('awardNamePlaceholder')}
                                fluid required maxLength={50}
                                value={this.state.awardName}
                                onChange={this.handleInputNameChange}
                            />
                        </div>
                    </div>
                    <div>
                        <Flex gap="gap.small">
                            <Text content={t('awardDescription')} size="medium" /><Text content="*" className="requiredfield" error size="medium" />
                            <Flex.Item push>
                                {this.getRequiredFieldError(this.state.isDescriptionValuePresent, t)}
                            </Flex.Item>
                        </Flex>
                        <div className="add-form-input">
                            <TextArea placeholder={t('awardDescriptionPlaceholder')}
                                fluid required maxLength={300}
                                className="response-text-area"
                                value={this.state.awardDescription}
                                onChange={this.handleInputDescriptionChange}
                            />
                        </div>
                    </div>
                    <div>
                        <Flex gap="gap.small">
                            <Text content={t('awardLink')} size="medium" />
                            <Flex.Item push>
                                <Text content={this.state.invalidLinkText} className="field-error-message" error size="medium" />
                            </Flex.Item>
                        </Flex>
                        <div>
                            <Input placeholder={t('awardLinkPlaceholder')} fluid required className="add-form-input"
                                value={this.state.awardImageLink}
                                onChange={this.handleInputImageChange}
                            />
                        </div>
                    </div>
                </div>
                <div className="tab-footer">
                    <div>
                        <Flex space="between">
                            <Button icon="icon-chevron-start"
                                content={t('backButtonText')} text
                                onClick={this.props.onBackButtonClick} />
                            <Flex gap="gap.small">
                                <Button content={t('addButtonText')} primary
                                    loading={this.state.isSubmitLoading}
                                    disabled={this.state.isSubmitLoading || !this.props.isNewAllowed}
                                    onClick={() => { this.onAddButtonClick(t) }}
                                />
                            </Flex>
                        </Flex>
                    </div>
                </div>
            </>
        );
    }
}

export default withTranslation()(AddAward)