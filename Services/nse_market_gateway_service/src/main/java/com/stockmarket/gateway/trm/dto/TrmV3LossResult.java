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
 *         TEJA RECOVERY MODEL (TRM) V3 LOSS SCENARIO RESULT DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Result payload for TRM V3 Path B (Loss Scenario with Calibrated Linear Recovery Surcharge).
 * Realizes a structural stop-loss exit, liquidates 100% of open equity, preserves unencumbered cash,
 * and calculates the subsequent recovery allocation:
 * <pre>
 *   StopPrice = EntryPrice * (1 - LossPercent)
 *   LossAmount = InitialInvestment * LossPercent
 *   PreservedCash = OriginalQuantity * StopPrice
 *   RecoverySurcharge = LossAmount * RecoveryFactor
 *   NextReversalAllocation = BaseReversalInvestment + RecoverySurcharge
 * </pre>
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok annotations.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Excludes null values.</li>
 *   <li><b>@JsonProperty</b>: Enforces consistent JSON keys.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Executed when an open trade triggers its structural stop-loss floor. Prevents martingale compounding
 * ruin by taking the loss immediately, preserving capital, and applying a minor recovery surcharge
 * to the next independent setup.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmV3LossResult implements Serializable {

    /**
     * Original entry purchase price per share (P_0).
     */
    @JsonProperty("entryPrice")
    private BigDecimal entryPrice;

    /**
     * Stop loss trigger price: EntryPrice * (1 - LossPercent).
     */
    @JsonProperty("stopPrice")
    private BigDecimal stopPrice;

    /**
     * Number of shares liquidated at the stop loss level.
     */
    @JsonProperty("originalQuantity")
    private BigDecimal originalQuantity;

    /**
     * Total seed cash invested: EntryPrice * OriginalQuantity.
     */
    @JsonProperty("initialInvestment")
    private BigDecimal initialInvestment;

    /**
     * Loss percentage incurred (e.g., 0.15 for 15%).
     */
    @JsonProperty("lossPercent")
    private BigDecimal lossPercent;

    /**
     * Realized dollar loss incurred: InitialInvestment * LossPercent.
     */
    @JsonProperty("lossAmount")
    private BigDecimal lossAmount;

    /**
     * Cash preserved upon 100% position liquidation: OriginalQuantity * StopPrice.
     */
    @JsonProperty("preservedCash")
    private BigDecimal preservedCash;

    /**
     * Linear surcharge factor applied to the realized loss (e.g., 0.15).
     */
    @JsonProperty("recoveryFactor")
    private BigDecimal recoveryFactor;

    /**
     * Absolute dollar surcharge added to the next trade: LossAmount * RecoveryFactor.
     */
    @JsonProperty("recoveryAmount")
    private BigDecimal recoveryAmount;

    /**
     * Standard baseline tranche capital for the next cycle (defaults to 100.00).
     */
    @JsonProperty("baseReversalInvestment")
    private BigDecimal baseReversalInvestment;

    /**
     * Next trade capital allocation: BaseReversalInvestment + RecoveryAmount.
     */
    @JsonProperty("nextReversalInvestment")
    private BigDecimal nextReversalInvestment;

    /**
     * Solvency check: true if NextReversalInvestment <= 15% of total bankroll equity.
     */
    @JsonProperty("bankrollCapPassed")
    private boolean bankrollCapPassed;

    /**
     * Action mandate instructing the broker/trader of risk preservation and next order sizing.
     */
    @JsonProperty("actionMandate")
    private String actionMandate;
}
