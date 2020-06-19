// <copyright file="manage-award-tab.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Menu, Loader } from "@fluentui/react-northstar";
import ManageAward from "./manage-awards";
import RewardCycle from "../award-cycle/reward-cycle";
import { SeverityLevel } from "@microsoft/applicationinsights-web";
import * as microsoftTeams from "@microsoft/teams-js";
import { createBrowserHistory } from "history";
import { WithTranslation, withTranslation } from "react-i18next";
import { getApplicationInsightsInstance } from "../../helpers/app-insights";
import "../../styles/site.css";
import { getAllAwards } from "../../api/awards-api";
const browserHistory = createBrowserHistory({ basename: "" });
interface IState {
    selectedMenuItemIndex: number,
    menuDisable: boolean,
    loader: boolean
}

/** Component for displaying on manage award tab. */
class AwardsTab extends React.Component<WithTranslation, IState> {
    telemetry?: any = null;
    locale?: string | null;
    theme?: string | null;
    teamId?: string | null;
    props: any;
    appInsights: any;
    userObjectId?: string = "";

    constructor(props) {
        super(props);
        this.state = {
            selectedMenuItemIndex: 0,
            menuDisable: false,
            loader: true
        };

        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.telemetry = params.get("telemetry");
    }

    getMenuItems = (t: any) => {
        return [
            {
                key: "manageawards",
                content: t('menuAwards'), 
            },
            {
                key: "rewardcycle",
                content: t('menuSetRewardCycle'),
                disabled: this.state.menuDisable
            }
        ];
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
            this.teamId = context.teamId;
            // Initialize application insights for logging events and errors.
            this.appInsights = getApplicationInsightsInstance(this.telemetry, browserHistory);
            this.getAwards();
        });
    }

    /**
   *Get awards from API
   */
    async getAwards() {
        this.appInsights.trackTrace({ message: `'getAwards' - Initiated request`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        let award = await getAllAwards(this.teamId!);
        if (award.data) {
            this.appInsights.trackTrace({ message: `'getAwards' - Request success`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
            let awards = award.data;
            
            if (awards.length === 0) {
                this.setState({ menuDisable: true });
            }
            else {
                this.setState({ menuDisable: false });
            }
        }
        else {
            this.appInsights.trackTrace({ message: `'getAwards' - Request failed`, properties: { User: this.userObjectId }, severityLevel: SeverityLevel.Information });
        }
        this.setState({
            loader: false
        });
    }

    /** 
   *  Called once menu item is clicked.
   * */
    onMenuItemClick = (event: any, data: any) => {
        this.setState({ selectedMenuItemIndex: data.index });
    }

    onModifyAward = (totalAwards: number) => {
        if (totalAwards === 0) {
            this.setState({ menuDisable: true })
        }
        else {
            this.setState({ menuDisable: false })
        }
    }

    render() {
        const { t } = this.props;
        if (this.state.loader) {
            return (
                <div className="module-container">
                    <Loader />
                </div>
            );
        } else {
            return (
                <div className="module-container">
                    <Menu defaultActiveIndex={0}  onItemClick={this.onMenuItemClick} items={this.getMenuItems(t)} className="manage-award-tab-menu"  underlined primary />
                    {this.state.selectedMenuItemIndex === 0 && <ManageAward teamId={this.teamId!} onModifyAward={this.onModifyAward} />}
                    {this.state.selectedMenuItemIndex === 1 && <RewardCycle teamId={this.teamId!} />}
                </div>
            );
        }
    }
}

export default withTranslation()(AwardsTab);