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
 *         TEJA RECOVERY MODEL (TRM) UNIFIED CALCULATION REQUEST DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Unified Data Transfer Object defining input simulation and live execution parameters across all
 * four Teja Recovery Model (TRM) algorithmic engines: Classic, V1, V2, and V3.
 *
 * Dual Mode Operation:
 * <ul>
 *   <li><b>Live Market Data Mode</b>: If {@code symbol} is supplied, missing prices, VWAP, and base
 *       capital are automatically enriched via live NSE exchange feeds by {@code TRMAlgorithmServiceImpl}.</li>
 *   <li><b>Synthetic Simulation Mode</b>: If explicit values are provided, the engine calculates pure
 *       hypothetical trajectories without requiring active market hours.</li>
 * </ul>
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data</b>: Lombok annotation generating getters, setters, {@code equals()}, {@code hashCode()}, and {@code toString()}.</li>
 *   <li><b>@Builder</b>: Generates a fluent builder pattern for constructing request instances.</li>
 *   <li><b>@NoArgsConstructor / @AllArgsConstructor</b>: Provides default and full-argument constructors required for Jackson and frameworks.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Excludes null fields from serialized JSON to minimize payload overhead.</li>
 *   <li><b>@JsonProperty</b>: Standardizes JSON key mapping across client interfaces.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Submitted via REST or synthesized by AI decision agents to compute risk thresholds, runner retention
 * levels, or scale-in limits before sending orders to execution brokers.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmCalculationRequest implements Serializable {

    /**
     * NSE equity ticker symbol (e.g. "RELIANCE", "TCS"). Triggers real-time market data enrichment.
     */
    @JsonProperty("symbol")
    private String symbol;

    /**
     * Reference purchase or entry price per share. Defaults to live market LTP if symbol is present.
     */
    @JsonProperty("entryPrice")
    private BigDecimal entryPrice;

    /**
     * Current traded price during an open trade. Defaults to entryPrice if not provided.
     */
    @JsonProperty("currentPrice")
    private BigDecimal currentPrice;

    /**
     * Initial position share volume (e.g. 100 shares).
     */
    @JsonProperty("quantity")
    private BigDecimal quantity;

    /**
     * Initial capital tranche allocated to Round 1 (S). Defaults to entryPrice * quantity.
     */
    @JsonProperty("baseCapital")
    private BigDecimal baseCapital;

    /**
     * Total account bankroll or trading fund equity, used to evaluate 15% risk caps.
     */
    @JsonProperty("bankroll")
    private BigDecimal bankroll;

    /**
     * Maximum number of recovery averaging rounds allowed (hard-capped at 5 in risk guidelines).
     */
    @JsonProperty("maxRounds")
    @Builder.Default
    private Integer maxRounds = 5;

    /**
     * Fixed geometric price drop percentage per round for TRM V1/V2 (e.g. 0.20 for 20% drawdown).
     */
    @JsonProperty("dropPercentage")
    @Builder.Default
    private BigDecimal dropPercentage = new BigDecimal("0.20");

    /**
     * Capital multiplier applied to previous tranche or cumulative investment (e.g. 2.0x).
     */
    @JsonProperty("multiplier")
    @Builder.Default
    private BigDecimal multiplier = new BigDecimal("2.0");

    /**
     * Target percentage gain for TRM V3 profit harvesting (e.g. 0.15 for 15% profit target).
     */
    @JsonProperty("profitTargetPercentage")
    @Builder.Default
    private BigDecimal profitTargetPercentage = new BigDecimal("0.15");

    /**
     * Stop loss threshold percentage for TRM V3 invalidation (e.g. 0.15 for 15% stop loss).
     */
    @JsonProperty("stopLossPercentage")
    @Builder.Default
    private BigDecimal stopLossPercentage = new BigDecimal("0.15");

    /**
     * Fraction of net profit converted into permanent zero-cost runner shares via Current Price Law (e.g. 0.15).
     */
    @JsonProperty("retentionFactor")
    @Builder.Default
    private BigDecimal retentionFactor = new BigDecimal("0.15");

    /**
     * Surcharge factor applied to realized dollar losses to calculate next trade recovery allocation (e.g. 0.15).
     */
    @JsonProperty("recoveryFactor")
    @Builder.Default
    private BigDecimal recoveryFactor = new BigDecimal("0.15");

    /**
     * Baseline tranche capital for the next cycle following a stop-loss liquidation (defaults to 100.00).
     */
    @JsonProperty("baseReversalInvestment")
    @Builder.Default
    private BigDecimal baseReversalInvestment = new BigDecimal("100.00");
}
