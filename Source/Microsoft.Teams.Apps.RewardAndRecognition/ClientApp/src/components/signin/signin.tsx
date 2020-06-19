import * as React from "react";
import { RouteComponentProps } from "react-router-dom";
import { useTranslation } from "react-i18next";
import { Button } from "@fluentui/react-northstar";
import * as microsoftTeams from "@microsoft/teams-js";

const SignInPage: React.FunctionComponent<RouteComponentProps> = props => {
    const { t } = useTranslation();
    function onSignIn() {
        microsoftTeams.initialize();
        microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/signin-simple-start",
            successCallback: () => {
                window.location.href = "/discover";
            },
            failureCallback: (reason) => {
                console.log("Login failed: " + reason);
                window.location.href = "/errorpage";
            }
        });
    }

    return (
        <div className="sign-in-content-container">
            <Button content={t('signInButtonText')} primary className="sign-in-button" onClick={onSignIn} />
        </div>
    );
};

export default SignInPage;