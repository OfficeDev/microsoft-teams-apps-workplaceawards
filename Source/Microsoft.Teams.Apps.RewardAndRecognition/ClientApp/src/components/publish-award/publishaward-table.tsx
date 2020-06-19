// <copyright file="publishaward-table.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Table, Text, Button, Accordion, Label, Dialog } from "@fluentui/react-northstar";
import CheckboxBase from "../checkbox-base";
import "../../styles/site.css";
import { useTranslation } from 'react-i18next';

interface IPublishAwardTableProps {
    showCheckbox: boolean,
    publishData: [],
    distinctAwards: [],
    onCheckBoxChecked: (nominationId: string, isChecked: boolean) => void,
    onChatButtonClick: (nominationDetails: any, t: any) => void
}

const PublishAwardTable: React.FunctionComponent<IPublishAwardTableProps> = props => {
    const { t } = useTranslation();
    const [dialogText, setDialogText] = React.useState("");
    const [dialogHeaderNominee, setDialogHeader] = React.useState(false);
    const [open, setOpen] = React.useState(false);
    const onItemClick = (content: string, isNominee: boolean) => {
        setDialogText(content);
        setDialogHeader(isNominee);
        setOpen(true);
    }
    const awardsTableHeader = {
        key: "header",
        items: props.showCheckbox === true ?
            [
                { content: < Text content={""} />, key: "check-box", className: "table-checkbox-cell" },
                { content: <Text weight="semibold" content={t('nomineesTableHeaderText')} />, className: "publish-table-nominee" },
                { content: <Text weight="semibold" content={t('nominationReasonTableHeaderText')} />, className: "publish-table-reason" },
                { content: <Text weight="semibold" content={t('endorsedByTableHeaderText')} />, className: "publish-table-endorse" },
                { content: <Text weight="semibold" content={t('chatWithNominatorTableHeaderText')} />, className: "publish-table-chat" }
            ]
            :
            [
                { content: <Text weight="semibold" content={t('publishResultHeaderText')} />, className: "publish-table-nominee-alluser" },
                { content: <Text weight="semibold" content={t('nominationReasonTableHeaderText')} />, className: "publish-table-reason-alluser" },
            ],
    };

    let awardsTableRows = props.publishData.map((value: any, index) => {
        let totalLabels = JSON.parse(value.GroupName);
        let labels = totalLabels.map((name, index) => {
            if (index <= 1) {
                return (
                    <Label className="label-color label-margin label-padding" onClick={() => onItemClick(value.NomineeNames, true)}  circular content={name} />
                )
            }
            else if (index === 2) {
                return (
                    <Label onClick={() => onItemClick(value.NomineeNames, true)} className="label-color label-padding" circular content={`+${totalLabels.length - 2}`} />
                )
            }
        }
        );


        return {
            key: value.AwardId,
            items: props.showCheckbox === true ?
                [
                    { content: <CheckboxBase onCheckboxChecked={props.onCheckBoxChecked} value={value.NominationId} />, key: index + "1", className: "table-checkbox-cell" },
                    { content: <>{labels}</>, key: index + "2", truncateContent: true, className: "publish-table-nominee" },
                    { content: <Text onClick={() => onItemClick(value.ReasonForNomination, false)} content={value.ReasonForNomination} title={value.ReasonForNomination} />, key: index + "3", truncateContent: true, className: "publish-table-reason" },
                    { content: <Text content={value.EndorsementCount} title={value.EndorsementCount} />, key: index + "4", truncateContent: true, className: "publish-table-endorse" },
                    {
                        content: <Button secondary onClick={() => props.onChatButtonClick(value, t)} title={value.NominatedByName} content={t('chatButtonText')} ></Button >, className: "publish-table-chat"
                    }
                ]
                :
                [
                    { content: <>{labels}</>, key: index + "2", truncateContent: true, className: "publish-table-nominee-alluser" },
                    { content: <Text className="word-break" onClick={() => onItemClick(value.ReasonForNomination, false)} content={value.ReasonForNomination} title={value.ReasonForNomination} />, key: index + "3", truncateContent: true, className: "publish-table-reason-alluser" },
                ],
        }
    });

    let panels = props.distinctAwards.map((value: any) => (
        {
            title: <Text content={value.AwardName} title={value.AwardName} weight="regular" className="award-header" />,
            content: <Table rows={awardsTableRows.filter(row => row.key === value.AwardId)} header={awardsTableHeader} className="table-cell-content" />
        }
    ));

    return (
        <div>
            <Accordion defaultActiveIndex={[0]} panels={panels} />
            <Dialog
                open={open}
                content={<Text className="word-break" content={dialogText} />}
                onCancel={() => setOpen(false)}
                cancelButton={t('buttonTextOk')}
                header={dialogHeaderNominee === true ? (props.showCheckbox ? t('nomineesTableHeaderText') : t('publishResultHeaderText')) : t('nominationReasonTableHeaderText')}
                headerAction={{
                    title: 'Close',
                    onClick: () => setOpen(false),
                }}
            />
        </div>
    );
}

export default PublishAwardTable;