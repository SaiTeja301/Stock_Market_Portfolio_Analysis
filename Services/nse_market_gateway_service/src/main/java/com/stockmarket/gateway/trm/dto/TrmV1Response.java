package com.stockmarket.gateway.trm.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.List;

/**
 * ====================================================================================================
 *         TEJA RECOVERY MODEL (TRM) V1 RESPONSE DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Response Data Transfer Object representing the output of the TRM V1 percentage-triggered stock equity model.
 * Governed by geometric price drop triggers: \(P_n = P_0 (1 - L)^n\) and capital sizing: \(B_{n+1} = m \cdot T_n\).
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok boilerplate.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Excludes null properties.</li>
 *   <li><b>@JsonProperty</b>: Standardizes JSON serialization field names.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Swing traders and long-term equity accumulation desks evaluate this model to calculate weighted average
 * price progression, aggregate share accumulation, and required recovery liquidity across staggered drops.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmV1Response implements Serializable {

    /**
     * Display name of the model ("TRM V1 - Percentage-Triggered Model").
     */
    @JsonProperty("modelName")
    @Builder.Default
    private String modelName = "TRM V1 - Percentage-Triggered Model";

    /**
     * Initial purchase price per share (P_0).
     */
    @JsonProperty("initialPrice")
    private BigDecimal initialPrice;

    /**
     * Fractional drop required to trigger subsequent rounds (e.g., 0.20 for 20%).
     */
    @JsonProperty("dropPercentage")
    private BigDecimal dropPercentage;

    /**
     * Capital multiplier applied to cumulative capital (e.g., 2.0x).
     */
    @JsonProperty("multiplier")
    private BigDecimal multiplier;

    /**
     * Total number of averaging drop rounds executed.
     */
    @JsonProperty("maxRounds")
    private int maxRounds;

    /**
     * Cumulative capital invested across all executed tranches.
     */
    @JsonProperty("totalInvested")
    private BigDecimal totalInvested;

    /**
     * Cumulative total shares accumulated across all executed tranches.
     */
    @JsonProperty("totalShares")
    private BigDecimal totalShares;

    /**
     * Current Weighted Average Price (WAP) of the accumulated position.
     */
    @JsonProperty("averageCost")
    private BigDecimal averageCost;

    /**
     * Terminal market price at the final simulated round.
     */
    @JsonProperty("currentPrice")
    private BigDecimal currentPrice;

    /**
     * Current aggregate market value of all held shares at terminal price.
     */
    @JsonProperty("marketValue")
    private BigDecimal marketValue;

    /**
     * Unrealized open mark-to-market Profit & Loss.
     */
    @JsonProperty("unrealizedPnL")
    private BigDecimal unrealizedPnL;

    /**
     * Projected capital allocation for the next round if prices fall further.
     */
    @JsonProperty("nextInvestment")
    private BigDecimal nextInvestment;

    /**
     * Granular step-by-step breakdown across each averaging round.
     */
    @JsonProperty("steps")
    private List<TrmTrancheItem> steps;

    /**
     * Diagnostic commentary regarding exponential capital risk and mathematical behavior.
     */
    @JsonProperty("commentary")
    private String commentary;
}
