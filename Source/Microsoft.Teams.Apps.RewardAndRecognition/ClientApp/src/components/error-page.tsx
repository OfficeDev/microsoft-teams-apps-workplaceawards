import * as React from "react";
import { Text, Flex, Provider, themes } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";
import { withTranslation, WithTranslation } from "react-i18next";
import "../styles/site.css";
import { Icon } from "office-ui-fabric-react";

class ErrorPage extends React.Component<WithTranslation> {
    code: string | null = null;
    message: string | null = null;

    constructor(props: any) {
        super(props);
        let search = window.location.search;
        let params = new URLSearchParams(search);
        this.code = params.get("code");
    }

    /** Called once component is mounted. */
    async componentDidMount() {
        microsoftTeams.initialize();
    }

    render() {
        const { t } = this.props;
        if (this.code === "401") {
            this.message = t('unauthorizedAccess');
        } else if (this.code === "403") {
            this.message = t('forbiddenErrorMessage');
        }
        else {
            this.message = t('errorMessage');
        }

        return (
            <div className="container-div">
                <Flex gap="gap.small" hAlign="center" vAlign="center" className="error-container">
                    <Flex gap="gap.small" hAlign="center" vAlign="center">
                        <Flex.Item>
                            <div className="error-div-align">
                                <Icon color="red" />
                            </div>
                        </Flex.Item>
                        <Flex.Item grow>
                            <Flex column gap="gap.small" vAlign="stretch">
                                <div>
                                    <Text weight="bold" error content={this.message} /><br />
                                </div>
                            </Flex>
                        </Flex.Item>
                    </Flex>
                </Flex>
            </div>
        );
    }
}

export default withTranslation()(ErrorPage);