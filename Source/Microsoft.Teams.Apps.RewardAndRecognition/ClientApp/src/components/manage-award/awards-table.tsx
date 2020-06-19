// <copyright file="awards-table.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Table, Text, Image } from "@fluentui/react-northstar";
import CheckboxBase from "../checkbox-base";
import { useTranslation } from 'react-i18next';
import "../../styles/site.css";
import { getBaseUrl } from "../../helpers/utility";

interface IAwardsTableProps {
    showCheckbox: boolean,
    awardsData: any[],
    onCheckBoxChecked: (awardId: string, isChecked: boolean) => void,
}

const AwardsTable: React.FunctionComponent<IAwardsTableProps> = props => {
    const { t } = useTranslation();
    const awardsTableHeader = {
        key: "header",
        items: props.showCheckbox === true ?
            [
                { content: <div />, key: "check-box", className: "table-checkbox-cell" },
                { content: <div />, key: "image", className: "table-image-cell" },
                {
                    content: <Text weight="semibold" content={t('awardName')} />, key: "response", className: "award-table-name"
                },
                { content: <Text weight="semibold" content={t('awardDescription')} />, key: "questions", className: "award-table-description" }
            ]
            :
            [
                { content: <Text weight="semibold" content={t('awardName')} />, key: "response", className: "award-table-name"},
                { content: <Text weight="semibold" content={t('awardDescription')} />, key: "questions", className: "award-table-description" }
            ],
    };

    let awardsTableRows = props.awardsData.map((value: any, index) => (
        {
            key: index,
            style: {},
            items: props.showCheckbox === true ?
                [
                    { content: <CheckboxBase onCheckboxChecked={props.onCheckBoxChecked} value={value.AwardId} />, key: index + "1", className: "table-checkbox-cell" },
                    {
                        content: <Image alt="NA" className="award-image-icon" fluid src={(value.awardLink === null || value.awardLink === "") ? getBaseUrl() + "/content/DefaultAwardImage.png" : value.awardLink} />
                        , key: index + "2", className: "table-image-cell"
                    },
                    { content: <Text content={value.AwardName} title={value.AwardName} />, key: index + "3", truncateContent: true, className: "award-table-name"},
                    { content: <Text content={value.awardDescription} title={value.awardDescription} />, key: index + "4", truncateContent: true, className:"award-table-description"}
                ]
                :
                [
                    { content: <Text content={value.AwardName} title={value.AwardName} />, key: index + "2", truncateContent: true, className: "award-table-name"},
                    { content: <Text content={value.awardDescription} title={value.awardDescription} />, key: index + "3", truncateContent: true, className:"award-table-description"}
                ],
        }
    ));

    return (
        <div>
            <Table rows={awardsTableRows}
                header={awardsTableHeader} className="table-cell-content" />
        </div>
    );
}

export default AwardsTable;