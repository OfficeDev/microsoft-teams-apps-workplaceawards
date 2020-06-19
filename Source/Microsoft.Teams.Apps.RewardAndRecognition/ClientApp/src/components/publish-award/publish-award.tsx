// <copyright file="publish-awards.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import { createBrowserHistory } from "history";
import { Button, Loader, Flex, Text, themes, Dialog, Icon, Image, Checkbox } from "@fluentui/react-northstar";
import { getAllAwardNominations, publishAwardNominations } from "../../api/nominate-awards-api";
import { getRewardCycle, setRewardCycle } from "../../api/reward-cycle-api";
import { sendWinnerNotification } from "../../api/notification-api";
import { getMembersInTeam, getUserRoleInTeam } from "../../api/configure-admin-api";
import { getBotSetting } from "../../api/setting-api";
import { getAllAwards } from "../../api/awards-api";
import PublishAwardTable from "./publishaward-table";
import ApprovedAwardTable from "./result-table";
import "../../styles/site.css";
import { RewardCycleState, RewardPublishState } from "../../models/award-cycle-state";
import Resources from "../../constants/resources";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import { ResultDetails, NominatedAward } from "../../models/result";
import { withTranslation, WithTranslation } from "react-i18next";
import { navigateToErrorPage, validateUserPartOfTeam } from "../../helpers/utility";
let moment = require('moment');

interface IState {
    Loader: boolean,
    isUserPartOfTeam: boolean,
    theme: string | null,
    themeStyle: any;
    errorMessage: string | null;
    selectedNominees: string[];
    publishAwardDataSet: any;
    distinctAwards: any;
    pubishResults: any;
    awardWinner: Array<ResultDetails>;
    activeAwardCycle: any;
    isNominationPriviewAvailable: boolean;
    openDialog: boolean;
    isWinnerCardSent: boolean;
    isAdminUser: boolean;
    isPublishedAwards: boolean;
    activeCycleId: string;
    isRewardCycleConfigured: boolean;
    currentAwardCycleDateRange: string | "";
    currentAwardCycleEndDate: string | "";
    distinctNominatedAwards: any;
    showPublishedWinners: boolean;
    showChatOption: boolean;
}

const browserHistory = createBrowserHistory({ basename: "" });

/** Component for displaying on publish award details. */
class PublishAward extends React.Component<WithTranslation, IState>
{
    locale?: string | null;
    telemetry?: any = null;
    appInsights: any;
    theme?: string | null;
    userEmail?: any = null;
    userObjectId?: string = "";
    teamId?: string | null;
    botId: string;
    appBaseUrl: string;
    appUrl: string = (new URL(window.location.href)).origin;


    constructor(props: any) {
        super(props);
        this.state = {
            Loader: true,
            isUserPartOfTeam: false,
            theme: this.theme ? this.theme : Resources.default,
            themeStyle: themes.teams,
            errorMessage: "",
            selectedNominees: [],
            publishAwardDataSet: [],
            distinctAwards: [],
            pubishResults: [],
            awardWinner: [],
            activeAwardCycle: {},
            isNominationPriviewAvailable: false,
            openDialog: false,
            isWinnerCardSent: false,
            isAdminUser: false,
            isPublishedAwards: false,
            activeCycleId: "",
            isRewardCycleConfigured: false,
            currentAwardCycleDateRange: "",
            distinctNominatedAwards: Array<NominatedAward>(),
            currentAwardCycleEndDate: "",
            showPublishedWinners: false,
            showChatOption: false
        };

        this.botId = '';
        this.appBaseUrl = window.location.origin;
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            this.userObjectId = context.userObjectId;
            this.userEmail = context.upn;
            this.teamId = context.teamId;
            this.theme = context.theme;
            this.locale = context.locale;
            this.getPageDetails();
        });       
    }

    /**
    * Get page details.
    */
    getPageDetails = async () => {
        let flag = await validateUserPartOfTeam(this.teamId!, this.userObjectId!)
        if (flag) {
            await this.getBotSetting(this.teamId!);
            await this.validateUserProfileInTeam();

            // get active award cycle details for admin user (isActivecycle: true), for non-admin user get recent most published award cycle details (isActivecycle: false).
            let flag = this.state.isAdminUser;            
            await this.getRewardCycle(flag);
            if (this.state.activeCycleId !== undefined || this.state.activeCycleId !== "") {
                await this.getPublishAwardDetails();
            }
        }
        else {
            navigateToErrorPage('');
        }
    }

    /**
   *Get bot id from API
   */
    async getBotSetting(teamId: string) {
        let response = await getBotSetting(teamId);
        if (response.data) {
            let settings = response.data;
            this.telemetry = settings.instrumentationKey;
            this.botId = settings.botId;

            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
        }
        else {
            navigateToErrorPage(response.status);
        }
    }

    submitHandler = async (err, result) => {
        this.appInsights.trackTrace(`Submit handler - err: ${err} - result: ${result}`);
        await this.updatePreviewState();
    };

    /**
    *Get award nomination details from API
    */
    async getRewardCycle(isActivecycle: boolean) {
        const { t } = this.props;
        this.appInsights.trackTrace({ message: `'getRewardCycle' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let response = await getRewardCycle(this.teamId!, isActivecycle)
        if (response.data) {
            this.appInsights.trackTrace({ message: `'getRewardCycle' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            let rewardcycle = response.data;
            this.setState({
                activeAwardCycle: rewardcycle,
                activeCycleId: rewardcycle.cycleId,
                isRewardCycleConfigured: true,
                currentAwardCycleDateRange: t('currentAwardCycleDateRange', { startDate: moment(rewardcycle.rewardCycleStartDate).format("ll"), endDate: moment(rewardcycle.rewardCycleEndDate).format("ll")})
            });
        }
        else {

            // check if reward cycle is configured by admin for non-admin preview
            let rewardResponse = await getRewardCycle(this.teamId!, true);
            if (rewardResponse.data) {
                this.setState({
                    isRewardCycleConfigured: true,
                    currentAwardCycleDateRange: t('currentAwardCycleDateRange', { startDate: moment(rewardResponse.data.rewardCycleStartDate).format("ll"), endDate: moment(rewardResponse.data.rewardCycleEndDate).format("ll")}),
                    currentAwardCycleEndDate: (moment(rewardResponse.data.rewardCycleEndDate).format("ll"))
                });
            }
            else {
                this.setState({
                    isRewardCycleConfigured: false,
                    currentAwardCycleDateRange: ""
                });
            }
        }
    }

    /**
    *Get award nomination details from API
    */
    async validateUserProfileInTeam() {
        this.appInsights.trackTrace({ message: `'getTeamMembersInTeam' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let teamMembers = await getMembersInTeam(this.teamId!);
        if (teamMembers.data) {
            this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });

            let member = teamMembers.data.find(element => element.aadobjectid === this.userObjectId);
            if (member !== null || member !== undefined) {
                this.setState({
                    isUserPartOfTeam: true
                });
                // check user role in team
                this.appInsights.trackTrace({ message: `'getUserRoleInTeam' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
                let adminDetails = await getUserRoleInTeam(this.teamId!);
                if (adminDetails.data) {
                    this.appInsights.trackTrace({ message: `'getUserRoleInTeam' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });

                    if (adminDetails.data.AdminObjectId === this.userObjectId) {
                        this.setState({
                            isAdminUser: true,
                            isPublishedAwards: false,
                            showChatOption: true
                        });
                    }
                    else {
                        this.setState({
                            isAdminUser: false,
                            isPublishedAwards: true,
                            showChatOption: false
                        });
                    }
                }
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'getTeamMembersInTeam' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            navigateToErrorPage(teamMembers.status);
        }
    }

    /**
    *Navigate to manage award tab
    */
    onManageAwardButtonClick = (t: any) => {
        microsoftTeams.tasks.startTask({
            completionBotId: this.botId,
            title: t('manageAwardButtonText'),
            height: 700,
            width: 700,
            url: `${this.appBaseUrl}/awards-tab?telemetry=${this.telemetry}&theme=${this.theme}&teamId=${this.teamId}&locale=${this.locale}`,
            fallbackUrl: `${this.appBaseUrl}/awards-tab?telemetry=${this.telemetry}&theme=${this.theme}&teamId=${this.teamId}&locale=${this.locale}`,
        }, this.submitHandler);
    }

    /**
    *Navigate to configure admin
    */
    onConfigureAdminButtonClick = (t: any) => {
        microsoftTeams.tasks.startTask({
            completionBotId: this.botId,
            title: t('configureAdminTitle'),
            height: 460,
            width: 600,
            url: `${this.appBaseUrl}/config-admin-page?telemetry=${this.telemetry}&teamId=${this.teamId}&theme=${this.theme}&locale=${this.locale}&updateAdmin=${true}`,
            fallbackUrl: `${this.appBaseUrl}/config-admin-page?telemetry=${this.telemetry}&teamId=${this.teamId}&theme=${this.theme}&locale=${this.locale}&updateAdmin=${true}`,
        }, this.submitHandler);
    }

    onChatButtonClick = (nominationDetails: any, t: any) => {
        let msg = t('chatWithNominatorMessageText', { nominator: nominationDetails.NominatedByName, nominees: nominationDetails.NomineeNames, award: nominationDetails.AwardName });
        let url = `https://teams.microsoft.com/l/chat/0/0?users=${nominationDetails.NominatedByUserPrincipalName}&message=${msg}`;
        microsoftTeams.executeDeepLink(url);
    }

    /**
    *Get selected nominations
    */
    onNominationSelected = (nominationId: string, isSelected: boolean) => {
        if (nominationId !== "") {
            let selectNominees = this.state.selectedNominees;
            let selectedAwardWinner = this.state.awardWinner;
            let nomination = this.state.publishAwardDataSet.filter(row => row.NominationId === nominationId).shift();
            if (isSelected) {
                selectNominees.push(nominationId);
                let results: ResultDetails = {
                    AwardId: nomination.AwardId,
                    AwardName: nomination.AwardName,
                    NominationId: nominationId,
                    WinnerCount: 0,
                    TeamId: this.teamId!,
                    NomineeNames: nomination.NomineeNames,
                    GroupName: nomination.GroupName,
                    NomineeObjectIds: nomination.NomineeObjectIds,
                    NomineeUserPrincipalNames: nomination.NomineeUserPrincipalNames,
                    AwardLink: this.state.distinctAwards.filter(row => row.AwardId === nomination.AwardId).shift() !== null ? this.state.distinctAwards.filter(row => row.AwardId === nomination.AwardId).shift().awardLink : "",
                    AwardCycle: this.state.currentAwardCycleDateRange,
                };
                selectedAwardWinner.push(results);
            }
            else {
                selectedAwardWinner.splice(selectNominees.indexOf(nominationId), 1);
                selectNominees.splice(selectNominees.indexOf(nominationId), 1);
            }

            this.setState({
                selectedNominees: selectNominees
            })

            this.setState({
                awardWinner: selectedAwardWinner
            })
        }
    }

    /**
    *Show publish award confirmation window
    */
    onPublishResultButtonClick = async (t: any) => {
        this.setState({
            selectedNominees: [],
            isRewardCycleConfigured: false,
            Loader: true,
        })
        let response = await this.publishAwards();
        if (response) {
            let winners = { TeamId: this.teamId, Winners: this.state.awardWinner };
            let notifyResponse = await sendWinnerNotification(winners);
            if (notifyResponse) {
                this.appInsights.trackEvent({ name: `Publish result` }, { User: this.userObjectId, Team: this.teamId! });

                // Update active award cycle to published
                let awardCycle = this.state.activeAwardCycle;
                awardCycle.resultPublished = RewardPublishState.Published;
                awardCycle.rewardCycleState = RewardCycleState.InActive;
                awardCycle.resultPublishedOn = new Date();
                let awardPublish = await setRewardCycle(awardCycle);
                if (awardPublish.data) {
                    this.setState({ openDialog: true })
                    this.setState({ isWinnerCardSent: true })
                }

                await this.updatePreviewState();
                await this.showWinnersInAdminTab();
            }
        }
        else {
            this.setState({ isWinnerCardSent: false })
        }
    }

    updatePublishState = async () => {
        this.setState({ openDialog: false, isNominationPriviewAvailable: false })
    }

    updatePreviewState = async () => {
        await this.validateUserProfileInTeam();
        await this.getRewardCycle(this.state.isAdminUser);
        await this.getPublishAwardDetails();
    }

    showWinnersInAdminTab = async () => {
        let flag = this.state.showPublishedWinners;
        this.setState({ showPublishedWinners: flag ? false : true, Loader: true })
        await this.getRewardCycle(flag);
        this.setState({ isPublishedAwards: !flag, showChatOption: flag })
        await this.getPublishAwardDetails();

        this.setState({
            Loader: false, selectedNominees: [], awardWinner: []
        });
    }

    /**
    *Publish award nominations from API
    */
    async publishAwards() {
        this.appInsights.trackTrace({ message: `'publishAwards' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let awards = await publishAwardNominations(this.teamId!, this.state.selectedNominees.toString());
        if (awards.data) {
            this.appInsights.trackTrace({ message: `'publishAwards' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            return true;
        }
        else {
            this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            return false;
        }
    }

    /**
    *Get award nomination details from API
    */
    async getPublishAwardDetails() {
        this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let nominations = await getAllAwardNominations(this.teamId!, this.state.isPublishedAwards, this.state.activeCycleId!);
        if (nominations.data) {
            this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            if (nominations.data.length > 0) {
                this.setState({
                    publishAwardDataSet: nominations.data,
                    isNominationPriviewAvailable: true
                });
            }
            else {
                this.setState({
                    isNominationPriviewAvailable: false
                });
            }

            this.appInsights.trackTrace({ message: `'getAwards' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            let awards = await getAllAwards(this.teamId!);
            if (awards.data) {
                this.appInsights.trackTrace({ message: `'getAllAwards' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });

                this.getNominatedAwards(nominations.data);

                this.setState({
                    distinctAwards: awards.data
                });
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'getPublishAwardDetails' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        this.setState({
            Loader: false
        });
    }

    getNominatedAwards = (nominations: any) => {
        let awardIds = nominations.map(item => item.AwardId)
            .filter((value, index, self) => self.indexOf(value) === index);

        this.setState({ distinctNominatedAwards: [] });
        let nominatedAwards = this.state.distinctNominatedAwards;
        console.log(nominations);

        awardIds.forEach(function (value: string | undefined) {
            let filter = nominations.filter(row => row.AwardId === value).shift();
            let results: NominatedAward = {
                AwardId: value,
                AwardName: filter.AwardName,
            };
            nominatedAwards.push(results);
        });

        nominatedAwards = nominatedAwards.sort(function (firstElement, secondElement) {
            var awardName = firstElement.AwardName.toLowerCase(), awardNameNext = secondElement.AwardName.toLowerCase();
            if (awardName < awardNameNext)
                return -1;
            if (awardName > awardNameNext)
                return 1;
            return 0;
        });

        this.setState({
            distinctNominatedAwards: nominatedAwards
        });
    }

    openPublishDialog = () => this.setState({ openDialog: true })
    closePublishDialog = () => this.setState({ openDialog: false })

    /**
   *Get wrapper page for selected awards for publish
   */
    private getPublishConfirmationPage = () => {
        if (!this.state.Loader) {
            return (
                <div>
                    <ApprovedAwardTable awardWinner={this.state.awardWinner}
                        distinctAwards={this.state.distinctNominatedAwards}
                    />
                </div>
            );
        }
    }

    private showWinnersTabToggle = (t: any) => {
        return (<Flex gap="gap.large" >
            {this.state.isAdminUser && <div><Text content={t('showWinnersInAdminTabText')} />
                <Checkbox toggle
                    checked={this.state.showPublishedWinners}
                    onChange={() => this.showWinnersInAdminTab()}
                /></div>
            }</Flex>);
    }

    private pageHeader = (t: any) => {
        return (<Flex gap="gap.small" wrap>
            {this.state.currentAwardCycleDateRange != "" &&
                <Text weight="semibold" className="award-cycle-header" content={(this.state.showChatOption ? t('rewardCycleText') : t('AwardWinnersCycleText')) + this.state.currentAwardCycleDateRange} />}
            {this.state.currentAwardCycleDateRange === "" && this.showWinnersTabToggle(t)}
            {this.state.isAdminUser &&
                <>
                <Flex.Item push>
                    <Flex>
                    <Button secondary className="publish-award-button" onClick={() => this.onConfigureAdminButtonClick(t)} content={t('configureAdminTitle')}></Button>
                    <Button className="publish-award-button" content={t('manageAwardButtonText')} onClick={() => this.onManageAwardButtonClick(t)} />
                    <Dialog
                        className="winner-dialog"
                        cancelButton={t('cancelButtonText')}
                        confirmButton={t('confirmButtonText')}
                        content={this.getPublishConfirmationPage()}
                        header={t('publishResultHeaderText')}
                        trigger={<Button primary disabled={this.state.selectedNominees.length === 0} content={t('grantAwardButtonText')}></Button>}
                        onConfirm={() => this.onPublishResultButtonClick(t)} />
                    {this.state.openDialog &&
                        <Dialog
                            className="winner-publish-dialog"
                            open={this.state.openDialog}
                            header={t('publishResultHeaderText')}
                            content={this.getPublishSuccessMessage(t)}
                            confirmButton={t('buttonTextOk')}
                            onConfirm={() => this.updatePublishState()}
                            headerAction={<Icon name="close" onClick={this.closePublishDialog} />}
                            />}
                        </Flex>
                    </Flex.Item>
                </>}
        </Flex>);
    }

    private getPublishSuccessMessage(t: any) {
        return (<Flex hAlign="center">
            <div>
                <div><Flex hAlign="center" vAlign="stretch">
                    <Image className="success-image" fluid src={this.state.isWinnerCardSent ? this.appUrl + "/content/SuccessIcon.png" : this.appUrl + "/content/ErrorIcon.png"} />
                </Flex></div>
                <div><Text weight="bold" content={this.state.isWinnerCardSent ? t('resultPublishSuccessMessage') : t('resultPublishFailedMessage')} /></div>
            </div>
        </Flex>)
    }

    /**
   *Get wrapper for page which acts as container for all child components
   */
    private getWrapperPage = (t: any) => {
        if (this.state.Loader) {
            return (
                <div className="loader">
                    <Loader />
                </div>
            );
        } else if (!this.state.Loader && this.state.isUserPartOfTeam && this.state.isNominationPriviewAvailable && this.state.isRewardCycleConfigured) {
            return (
                <div>
                    <PublishAwardTable showCheckbox={this.state.showChatOption}
                        publishData={this.state.publishAwardDataSet}
                        distinctAwards={this.state.distinctNominatedAwards}
                        onCheckBoxChecked={this.onNominationSelected}
                        onChatButtonClick={this.onChatButtonClick}
                    />
                </div>
            );
        }
        else if (!this.state.Loader && !this.state.isNominationPriviewAvailable && this.state.isRewardCycleConfigured && this.state.isAdminUser) {
            return (<Flex className="error-container" hAlign="center" vAlign="stretch">
                <div>
                    <div><Flex hAlign="center" vAlign="stretch">
                        <Image className="preview-image" fluid src={this.appUrl + "/content/messages.png"} />
                    </Flex></div>
                    <div><Text content={!this.state.showPublishedWinners ? t('nominationPreviewMessage') : (t('resultPublishNotificationText') + this.state.currentAwardCycleEndDate)} /></div>
                </div>
            </Flex>)
        }
        else if (!this.state.Loader && !this.state.isRewardCycleConfigured) {
            return (<Flex className="error-container" hAlign="center" vAlign="stretch">
                <div>
                    <div><Flex hAlign="center" vAlign="stretch">
                        <Image className="preview-image" fluid src={this.appUrl + "/content/messages.png"} />
                    </Flex></div>
                    <div><Text content={t('cycleValidationMessage')} /></div>
                </div>
            </Flex>)
        }
        else if (!this.state.Loader && this.state.isRewardCycleConfigured && this.state.currentAwardCycleDateRange !== "" && !this.state.isAdminUser) {
            return (<Flex className="error-container" hAlign="center" vAlign="stretch">
                <div>
                    <div><Flex hAlign="center" vAlign="stretch">
                        <Image className="preview-image" fluid src={this.appUrl + "/content/messages.png"} />
                    </Flex></div>
                    <div><Text content={t('resultPublishNotificationText') + this.state.currentAwardCycleEndDate} /></div>
                </div>
            </Flex>)
        }
    }

    /**
  * Renders the component
  */
    public render() {
        const { t } = this.props;
        return (
            <div className="page-container">
                <div className="publish-table-container">
                    {this.pageHeader(t)}
                    {this.state.currentAwardCycleDateRange != "" && this.showWinnersTabToggle(t)}
                    <div>
                        {this.getWrapperPage(t)}
                    </div>
                </div>
            </div>
        );
    }
}

export default withTranslation()(PublishAward);