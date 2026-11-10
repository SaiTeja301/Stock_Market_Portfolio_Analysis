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
 *         TEJA RECOVERY MODEL (TRM) V2 ADAPTIVE RESPONSE DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Response Data Transfer Object representing the output of the TRM V2 adaptive actual-portfolio-loss model.
 * Sizing is strictly governed by open mark-to-market drawdown:
 * <pre>
 *   MV = Q_{total} * P_{curr}
 *   L_{actual} = max(0, T - MV)
 *   B_{next} = m * L_{actual}
 * </pre>
 * Delivers ~99.68% capital savings compared to TRM V1 while neutralizing position cost basis.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok annotations.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Excludes null values.</li>
 *   <li><b>@JsonProperty</b>: Enforces standard JSON serialization properties.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Utilized by institutional risk desks and capital-constrained retail traders to calculate precise
 * scale-in tranches that neutralize loss without escalating exposure beyond the 15% bankroll cap.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmV2Response implements Serializable {

    /**
     * Display name of the model ("TRM V2 - Actual-Loss Adaptive Model").
     */
    @JsonProperty("modelName")
    @Builder.Default
    private String modelName = "TRM V2 - Actual-Loss Adaptive Model";

    /**
     * Initial entry price per share (P_0).
     */
    @JsonProperty("initialPrice")
    private BigDecimal initialPrice;

    /**
     * Current stock price during the evaluated drawdown stage.
     */
    @JsonProperty("currentPrice")
    private BigDecimal currentPrice;

    /**
     * Multiplier applied to actual open mark-to-market loss (typically 2.0).
     */
    @JsonProperty("lossMultiplier")
    private BigDecimal lossMultiplier;

    /**
     * Current aggregate market value of all held shares: TotalShares * CurrentPrice.
     */
    @JsonProperty("marketValue")
    private BigDecimal marketValue;

    /**
     * Actual open mark-to-market dollar loss: max(0, TotalInvested - MarketValue).
     */
    @JsonProperty("actualLoss")
    private BigDecimal actualLoss;

    /**
     * Next tranche capital requirement: LossMultiplier * ActualLoss.
     */
    @JsonProperty("nextInvestment")
    private BigDecimal nextInvestment;

    /**
     * Number of additional shares to purchase in the next tranche: NextInvestment / CurrentPrice.
     */
    @JsonProperty("additionalShares")
    private BigDecimal additionalShares;

    /**
     * Total capital invested across all prior tranches plus the next planned tranche.
     */
    @JsonProperty("totalInvested")
    private BigDecimal totalInvested;

    /**
     * Total shares accumulated across all prior tranches plus the next planned tranche.
     */
    @JsonProperty("totalShares")
    private BigDecimal totalShares;

    /**
     * New Weighted Average Price (WAP) after executing the next recovery tranche.
     */
    @JsonProperty("averageCost")
    private BigDecimal averageCost;

    /**
     * Open unrealized Profit & Loss.
     */
    @JsonProperty("unrealizedPnL")
    private BigDecimal unrealizedPnL;

    /**
     * Target exit price required to realize a +5% net recovery gain above the new average cost.
     */
    @JsonProperty("recoveryTargetPrice")
    private BigDecimal recoveryTargetPrice;

    /**
     * Percentage capital saved by TRM V2 relative to TRM V1 for identical price drops (~99.68%).
     */
    @JsonProperty("capitalSavingsVsV1Percent")
    private BigDecimal capitalSavingsVsV1Percent;

    /**
     * True if cumulative capital deployed exceeds the 15% single-stock bankroll limit.
     */
    @JsonProperty("bankrollCapBreached")
    private boolean bankrollCapBreached;

    /**
     * True if the number of averaging rounds exceeds the maximum limit of 5.
     */
    @JsonProperty("maxRoundsCapBreached")
    private boolean maxRoundsCapBreached;

    /**
     * Granular per-round simulation step progression.
     */
    @JsonProperty("steps")
    private List<TrmTrancheItem> steps;

    /**
     * Executive action directive (e.g., "DEPLOY TRANCHE", "HALT: Bankroll Cap Breached").
     */
    @JsonProperty("actionDirective")
    private String actionDirective;
}
