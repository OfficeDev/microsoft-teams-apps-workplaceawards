/*
    <copyright file="configure-admin.tsx" company="Microsoft">
    Copyright (c) Microsoft. All rights reserved.
    </copyright>
*/

import * as React from "react";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import { createBrowserHistory } from "history";
import { Dropdown, Button, Loader, Flex, Text, themes, TextArea } from "@fluentui/react-northstar";
import { saveAdminDetails, getMembersInTeam } from "../../api/configure-admin-api";
import Constants from "../../constants/constants";
import { withTranslation, WithTranslation } from "react-i18next";
import "../../styles/site.css";
import { AdminDetails } from "../../models/admin-detail";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { navigateToErrorPage, validateUserPartOfTeam } from "../../helpers/utility";

interface IState {
    loading: boolean,
    theme: string | null,
    themeStyle: any;
    noteForTeam: string;
    allMembers: any[];
    selectedMember: any;
    isSubmitLoading: boolean;
    isSelectedMemberPresent: boolean;
    errorMessage: string | null;
    searchQuery: any;
}

const browserHistory = createBrowserHistory({ basename: "" });

/** Component for displaying on admin details. */
class ConfigurationAdminPage extends React.Component<WithTranslation, IState>
{
    locale?: string | null;
    telemetry?: any = null;
    appInsights: any;
    theme?: string | null;
    userEmail?: any = null;
    userObjectId?: string | null = null;
    teamId?: string | null;
    updateAdmin?: string | null;

    constructor(props: any) {
        super(props);
        this.state = {
            loading: false,
            theme: null,
            themeStyle: themes.teams,
            noteForTeam: "",
            allMembers: [],
            selectedMember: null,
            isSubmitLoading: false,
            isSelectedMemberPresent: true,
            errorMessage: "",
            searchQuery: undefined,
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
        this.updateAdmin = params.get("updateAdmin");
    }

    /** 
     *  Called once component is mounted. 
    */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.userEmail = context.upn;
            this.theme = context.theme;
            this.locale = context.locale;
            this.teamId = context.teamId;

            // Initialize application insights for logging events and errors.
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
            let flag = validateUserPartOfTeam(this.teamId!, this.userObjectId!)
            if (flag) {
                this.getMembersInTeam();
            }
            else {
                navigateToErrorPage('');
            }
        });
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
    }

    /**
     * Handle save admin detail event.
    */
    SaveAdminDetails = async (t: any) => {
        let selectedMember = this.state.selectedMember;
        if (selectedMember === null) {
            this.setState({ isSelectedMemberPresent: false });
            return;
        }
        this.setState({ isSubmitLoading: true });
        let admin = this.state.selectedMember;
        let member = this.state.allMembers.find(element => element.aadobjectid === this.userObjectId);
        let adminDetails: AdminDetails =
        {
            CreatedByUserPrincipalName: member.content,
            CreatedByObjectId: this.userObjectId != null ? this.userObjectId : null,
            CreatedOn: new Date(),
            AdminName: admin.header,
            AdminObjectId: admin.aadobjectid,
            AdminUserPrincipalName: admin.content,
            NoteForTeam: this.state.noteForTeam,
            TeamId: this.teamId!
        };

        this.appInsights.trackTrace({ message: `'saveAdminDetails' - Request initiated`, severityLevel: SeverityLevel.Information, properties: { UserEmail: this.userEmail } });
        const saveAdminDetailsResponse = await saveAdminDetails(adminDetails);
        if (saveAdminDetailsResponse.data) {
            this.appInsights.trackTrace({ message: `'saveAdminDetails' - Request success`, severityLevel: SeverityLevel.Information, properties: { UserEmail: this.userEmail } });
            this.appInsights.trackEvent({ name: `Set Champion` }, { User: this.userObjectId, Team: this.teamId! });
            let toBot =
            {
                Command: this.updateAdmin ?.toUpperCase() === "TRUE" ? Constants.SaveAdminDetailCommand : Constants.UpdateAdminDetailCommand,
                AdminName: admin.header,
                AdminUserPrincipalName: admin.content,
                NoteForTeam: this.state.noteForTeam,
                TeamId: this.teamId!
            };

            microsoftTeams.tasks.submitTask(toBot);
        }
        else {
            this.setState({ isSubmitLoading: false, errorMessage: t('errorMessage') });
            this.appInsights.trackTrace({ message: `'saveAdminDetails' - Request failed`, severityLevel: SeverityLevel.Error, properties: { UserEmail: this.userEmail, Code: saveAdminDetailsResponse.status } });
        }
    }

    Cancel = async () => {
        let toBot = { Command: Constants.CancelCommand };
        microsoftTeams.tasks.submitTask(toBot);
    }

    onNoteChange(event) {
        this.setState({
            noteForTeam: event.target.value
        });
    }

    getA11SelectionMessage = {
        onAdd: item => {
            if (item) {
                this.setState({ selectedMember: item, isSelectedMemberPresent: true });
            }
            return "";
        }
    };

    onSearchQueryChange = (event: any, props: any) => {
        if (props.value) {
            let searchQuery = props.searchQuery;
            let selectedItem = props.items.find((item) => item.header.toUpperCase() === searchQuery.toUpperCase());
            if (selectedItem) {
                this.setState({ selectedMember: selectedItem, isSelectedMemberPresent: true, searchQuery: searchQuery });
            } else {
                this.setState({ selectedMember: null, isSelectedMemberPresent: false, searchQuery: undefined });
            }
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

    render() {
        if (this.state.loading) {
            return (
                <div className="loader">
                    <Loader />
                </div>
            );
        } else {
            const { t } = this.props;
            return (
                <div className="add-user-responses-page">
                    <Flex gap="gap.smaller" vAlign="center" >
                        <Text content={t('selectTeamMemberTitle')} /><Text content="*" className="requiredfield" error size="medium" />
                        <Flex.Item push>
                            {this.getRequiredFieldError(this.state.isSelectedMemberPresent, t)}
                        </Flex.Item>
                    </Flex>
                    <Flex vAlign="center">
                        <Flex.Item align="start" grow>
                            <Dropdown className="add-form-input"
                                fluid
                                search
                                items={this.state.allMembers}
                                placeholder={t('selectTeamMemberPlaceHolder')}
                                getA11ySelectionMessage={this.getA11SelectionMessage}
                                noResultsMessage={t('noMatchesFoundText')}
                                onSearchQueryChange={this.onSearchQueryChange}
                                searchQuery={this.state.searchQuery}
                                value={this.state.selectedMember}
                            />
                        </Flex.Item>
                    </Flex>
                    <div>
                        <Flex vAlign="center">
                            <Text content={t('noteForTeamTitle')} />
                        </Flex>
                        <div className="add-form-input">
                            <TextArea fluid
                                maxLength={200}
                                className="response-text-area"
                                placeholder={t('noteForTeamPlaceHolder')}
                                value={this.state.noteForTeam}
                                onChange={this.onNoteChange.bind(this)}
                            />
                        </div>
                    </div>
                    <div className="error">
                        <Flex gap="gap.small">
                            {this.state.errorMessage !== null && <Text className="small-margin-left" content={this.state.errorMessage} error />}
                        </Flex>
                    </div>
                    <div className="tab-footer">
                        <Flex gap="gap.small" hAlign="end">
                            <Button secondary content={t('cancelButtonText')} onClick={this.Cancel} />
                            <Button primary content={t('saveButtonText')} loading={this.state.isSubmitLoading} disabled={this.state.isSubmitLoading} onClick={() => { this.SaveAdminDetails(t) }} />
                        </Flex>
                    </div>
                </div>
            );
        }
    }
}

export default withTranslation()(ConfigurationAdminPage)
