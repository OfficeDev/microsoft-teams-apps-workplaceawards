// <copyright file="result-table.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import { Text, Flex, Grid } from "@fluentui/react-northstar";
import "../../styles/site.css";
import { useTranslation } from 'react-i18next';

interface IApprovedAwardTableProps {
    awardWinner: any,
    distinctAwards: any
}

const ApprovedAwardTable: React.FunctionComponent<IApprovedAwardTableProps> = props => {
    const { t } = useTranslation();

    let gridItems = props.distinctAwards.map((value: any, index) => {
        let winnerCount = props.awardWinner.filter(a => a.AwardId === value.AwardId).length;
        return (
            <Flex column padding="padding.medium">
                <Text weight="semibold" content={value.AwardName} className="word-break" />
                <Text weight="bold" content={winnerCount > 1 ? winnerCount + " " + t('winnersCountText') : winnerCount + " " + t('winnerCountText')} />
            </Flex>
        )
    });

    return (
        <div className="result-container">
            <Grid columns="2" content={gridItems} />
        </div>
    );
}

export default ApprovedAwardTable;