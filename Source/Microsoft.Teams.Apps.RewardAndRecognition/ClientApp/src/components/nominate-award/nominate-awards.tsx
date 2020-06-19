/*
    <copyright file="nominate-awards.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import { createBrowserHistory } from "history";
import { Dropdown, Button, Loader, Flex, Text, themes, TextArea } from "@fluentui/react-northstar";
import { getMembersInTeam } from "../../api/configure-admin-api";
import "../../styles/site.css";
import { withTranslation, WithTranslation } from "react-i18next";
import { getAllAwards } from "../../api/awards-api";
import { NominationAwardPreview } from "../../models/nomination-award-preview";
import PreviewAward from "./preview-nominated-award";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { navigateToErrorPage, isNullorWhiteSpace } from "../../helpers/utility";

interface IState {
    loading: boolean,
    theme?: string | null,
    themeStyle: any;
    reasonForNomination: string;
    allMembers: any[];
    awards: any[];
    selectedAward: any;
    selectedMembers: any[];
    isSubmitLoading: boolean;
    isSelectedMembersPresent: boolean;
    isSelectedAwardPresent: boolean;
    errorMessage: string | null;
    isPreviewAward: boolean;
    awardDescription: string;
    isNoteForNomination: boolean;
}

const browserHistory = createBrowserHistory({ basename: "" });

/** Component for displaying on nomination details. */
class NominateAwards extends React.Component<WithTranslation, IState>
{
    locale?: string | null;
    awardId?: string | null;
    telemetry?: any = null;
    appInsights: any;
    theme?: string | null | undefined;
    userEmail?: any = null;
    userObjectId?: string | null = null;
    teamId?: string | null;
    appUrl: string = (new URL(window.location.href)).origin;

    constructor(props: any) {
        super(props);
        this.state = {
            loading: true,
            theme: null,
            themeStyle: themes.teams,
            reasonForNomination: "",
            allMembers: [],
            selectedAward: null,
            awards: [],
            selectedMembers: [],
            isSubmitLoading: false,
            isSelectedMembersPresent: true,
            isSelectedAwardPresent: true,
            errorMessage: "",
            isPreviewAward: false,
            awardDescription: "",
            isNoteForNomination: true,
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.awardId = params.get("awardId")!;
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.userEmail = context.upn;
            this.theme = context.theme;
            this.locale = context.locale;
            this.teamId = context.teamId;
            this.getPageDetails();
        });
    }

    /**
    * Get page details.
    */
    getPageDetails = async () => {
        // Initialize application insights for logging events and errors.
        this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);

        await this.getAwards();
        await this.getMembersInTeam();
        if (this.awardId != null) {

            let award = this.state.awards.find(element => element.key === this.awardId);
            this.setState({ selectedAward: award, awardDescription: award.description })
            this.setState({ loading: false });
        }
    }

    /**
    *  Get awards from API
    */
    async getAwards() {
        this.appInsights.trackTrace({ message: `'getAwards' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let awards = await getAllAwards(this.teamId!);
        if (awards.data) {
            this.appInsights.trackTrace({ message: `'getAwards' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });

            let awardDetails: any[] = [];

            awards.data.forEach((award) => {
                awardDetails.push({
                    key: award.AwardId,
                    header: award.AwardName,
                    imageUrl: award.awardLink === "" ? this.appUrl + "/content/DefaultAwardImage.png" : award.awardLink,
                    description: award.awardDescription,
                });
            });
            this.setState({
                awards: awardDetails,
            });
        }
        else {
            this.appInsights.trackTrace({ message: `'getAwards' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            this.setState({ loading: false });
            navigateToErrorPage(awards.status);
        }
    }

    /** 
    *  Get all team members.
    */
    getMembersInTeam = async () => {
        this.appInsights.trackTrace({ message: `'getMembersInTeam' - Request initiated`, severityLevel: SeverityLevel.Information });
        
        const teamMemberResponse = await getMembersInTeam(this.teamId!);
        if (teamMemberResponse) {
            if (teamMemberResponse.data) {
                this.setState({ allMembers: teamMemberResponse.data });
            }
            else {
                this.appInsights.trackTrace({ message: `'getMembersInTeam' - Request failed:${teamMemberResponse.status}`, severityLevel: SeverityLevel.Error, properties: { Code: teamMemberResponse.status } });
                navigateToErrorPage(teamMemberResponse.status);
            }
        }

        this.setState({ loading: false });
    }

    /**
       * Triggers when user clicks back button
    */
    onBackButtonClick = () => {
        this.setState({ isPreviewAward: false, isSubmitLoading: false });
    }

    /**
    *Navigate to nomination preview page
    */
    onPreviewButtonClick = () => {
        this.appInsights.trackEvent({ name: `Preview award` }, { User: this.userObjectId, Team: this.teamId! });
        if (this.state.selectedMembers.length === 0) {
            this.setState({ isSelectedMembersPresent: false });
        }
        if (this.state.selectedAward === null) {
            this.setState({ isSelectedAwardPresent: false });
        }
        if (isNullorWhiteSpace(this.state.reasonForNomination)) {
            this.setState({ isNoteForNomination: false });

            return;
        }
        if (this.state.selectedMembers.length > 0 && this.state.selectedAward != null) {
            this.setState({ isPreviewAward: true, isSubmitLoading: true });
        }        
    }

    /**
    *  Returns layout for preview nominated award.
    * */
    showNominatedAwardPreview = (): JSX.Element | undefined => {
        let recipients: Array<any> = [];
        this.state.selectedMembers.forEach((value) => {
            recipients.push(value.header);
        });
        let userPrincipalId: Array<any> = [];
        this.state.selectedMembers.forEach((value) => {
            userPrincipalId.push(value.content);
        });
        let member = this.state.allMembers.find(element => element.aadobjectid === this.userObjectId);
        let objectId: Array<any> = [];
        this.state.selectedMembers.forEach((value) => {
            objectId.push(value.aadobjectid);
        });
        let award = this.state.selectedAward;
        let nominatedAwardDetails: NominationAwardPreview =
        {
            AwardName: award.header,
            AwardId: award.key,
            ImageUrl: award.imageUrl,
            ReasonForNomination: this.state.reasonForNomination,
            Nominees: recipients,
            TeamId: this.teamId!,
            NominatedByName: member.header,
            NominatedByObjectId: this.userObjectId,
            NominatedByUserPrincipalName: this.userEmail,
            NomineeUserPrincipalNames: userPrincipalId,
            NomineeObjectIds: objectId,
            telemetry: this.telemetry,
            theme: this.theme,
            locale: this.locale,
        };

        return (
            <PreviewAward NominationAwardPreview={nominatedAwardDetails} onBackButtonClick={this.onBackButtonClick} />
        );
    }

    onNoteChange(event) {
        this.setState({
            reasonForNomination: event.target.value,
            isNoteForNomination: true
        });
    }

    getA11SelectionMessage = {
        onAdd: item => {
            let selectedMembers = this.state.selectedMembers;
            selectedMembers.push(item);
            this.setState({ selectedMembers: selectedMembers });
            if (this.state.selectedMembers.length > 0) { this.setState({ isSelectedMembersPresent: true }); }

            return "";
        },

        onRemove: item => {
            let selectedMembers = this.state.selectedMembers;
            selectedMembers.splice(selectedMembers.indexOf(item), 1);
            this.setState({ selectedMembers: selectedMembers });

            return "";
        }
    };

    onAwardsSelected = {
        onAdd: item => {
            if (item) {
                let award = this.state.awards.find(element => element.key === item.key)
                this.setState({ selectedAward: award, awardDescription: award.description, isSelectedAwardPresent: true });    
            }
            
            return "";
        }
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

    renderNominateAwards() {
        const { t } = this.props;
        return (
            <div className="container-subdiv-main">
                <Flex gap="gap.large" vAlign="center" className="title">
                    <Text content={t('selectAwardTitle')} />
                    <Flex.Item push>
                        {this.getRequiredFieldError(this.state.isSelectedAwardPresent, t)}
                    </Flex.Item>
                </Flex>
                <Flex gap="gap.large" vAlign="center">
                    <Flex.Item align="start" size="size.small" grow>
                        <Dropdown
                            fluid
                            items={this.state.awards}
                            placeholder={t('selectAwardPlaceHolder')}
                            getA11ySelectionMessage={this.onAwardsSelected}
                            noResultsMessage={t('noMatchesFoundText')}
                            value={this.state.selectedAward}
                        />
                    </Flex.Item>
                </Flex>
                <div>
                    <Flex gap="gap.large" vAlign="center" className="title">
                        <Text content={t('awardDescription')} />
                    </Flex>
                    <Text className="response-text-area word-break"
                        content={this.state.awardDescription}
                    />
                </div>
                <Flex gap="gap.small" vAlign="center" className="title">
                    <Text content={t('toBeAwardedToTitle')} /><Text content="*" className="requiredfield" error size="medium" />
                    <Flex.Item push>
                        {this.getRequiredFieldError(this.state.isSelectedMembersPresent, t)}
                    </Flex.Item>
                </Flex>
                <Flex gap="gap.large" vAlign="center">
                    <Flex.Item align="start" size="size.small" grow>
                        <Dropdown
                            fluid
                            search
                            multiple
                            items={this.state.allMembers}
                            placeholder={t('toBeAwardedToPlaceHolder')}
                            getA11ySelectionMessage={this.getA11SelectionMessage}
                            noResultsMessage={t('noMatchesFoundText')}
                            value={this.state.selectedMembers}
                        />
                    </Flex.Item>
                </Flex>
                <div>
                    <Flex gap="gap.small" vAlign="center" className="title">
                        <Text content={t('reasonForNominationTitle')} /><Text content="*" className="requiredfield" error size="medium" />
                        <Flex.Item push>
                            {this.getRequiredFieldError(this.state.isNoteForNomination, t)}
                        </Flex.Item>
                    </Flex>
                    <TextArea fluid
                        maxLength={300}
                        className="reasonfornomination-text-area"
                        placeholder={t('reasonForNominationPlaceHolder')}
                        value={this.state.reasonForNomination}
                        onChange={this.onNoteChange.bind(this)}
                    />
                </div>
                <div className="error">
                    <Flex gap="gap.small">
                        {this.state.errorMessage !== null && <Text className="small-margin-left" content={this.state.errorMessage} error />}
                    </Flex>
                </div>
                <div className="tab-footer preview-button">
                    <Flex gap="gap.small" hAlign="end">
                        <Button primary content={t('previewButtonText')} loading={this.state.isSubmitLoading} disabled={this.state.isSubmitLoading} onClick={this.onPreviewButtonClick} />
                    </Flex>
                </div>
            </div>
        );
    }

    render() {
        let contents = this.state.isPreviewAward
            ? this.showNominatedAwardPreview()
            : this.renderNominateAwards();

        if (this.state.loading) {
            return (
                <div>
                    <Loader />
                </div>
            );
        }
        else {
            return (
                <div className="module-container">
                    {contents}
                </div>
            );
        }
    }
}

export default withTranslation()(NominateAwards)
