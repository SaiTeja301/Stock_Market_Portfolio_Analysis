package com.stockmarket.gateway.trm.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.math.BigDecimal;

/**
 * ====================================================================================================
 *         TEJA RECOVERY MODEL (TRM) TRANCHE STEP ITEM DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Represents the detailed execution state of an individual round/tranche within a multi-round
 * TRM accumulation or recovery progression (used across Classic, V1, and V2).
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok boilerplate.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Minimizes JSON payload size.</li>
 *   <li><b>@JsonProperty</b>: Guarantees standardized serialization keys.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Rendered in UI tranche execution tables, instructing the trading desk or automated order router
 * of the exact limit price to bid, shares to acquire, and resulting weighted average price (WAP).
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmTrancheItem implements Serializable {

    /**
     * Sequential round index (e.g. 1 for base entry, 2 for first dip tranche).
     */
    @JsonProperty("round")
    private int round;

    /**
     * Stock market price triggering or executing this tranche.
     */
    @JsonProperty("stockPrice")
    private BigDecimal stockPrice;

    /**
     * Dollar/rupee capital allocated specifically to this tranche.
     */
    @JsonProperty("trancheInvestment")
    private BigDecimal trancheInvestment;

    /**
     * Number of shares acquired at this tranche's execution price.
     */
    @JsonProperty("sharesBought")
    private BigDecimal sharesBought;

    /**
     * Cumulative capital deployed across all rounds up to and including this one.
     */
    @JsonProperty("totalInvested")
    private BigDecimal totalInvested;

    /**
     * Cumulative total shares accumulated across all rounds.
     */
    @JsonProperty("totalShares")
    private BigDecimal totalShares;

    /**
     * Effective Weighted Average Price (WAP) of the entire accumulated position.
     */
    @JsonProperty("averageCost")
    private BigDecimal averageCost;

    /**
     * Current aggregate market valuation of the total accumulated shares.
     */
    @JsonProperty("marketValue")
    private BigDecimal marketValue;

    /**
     * Open unrealized Profit & Loss (Market Value - Total Invested).
     */
    @JsonProperty("unrealizedPnL")
    private BigDecimal unrealizedPnL;

    /**
     * Actual open mark-to-market loss: max(0, Total Invested - Market Value).
     */
    @JsonProperty("actualLoss")
    private BigDecimal actualLoss;

    /**
     * Projected capital requirement for the subsequent tranche if prices continue falling.
     */
    @JsonProperty("nextInvestment")
    private BigDecimal nextInvestment;

    /**
     * Calculated target price required to achieve profitable position exit (+5% above average cost).
     */
    @JsonProperty("targetExitPrice")
    private BigDecimal targetExitPrice;
}
