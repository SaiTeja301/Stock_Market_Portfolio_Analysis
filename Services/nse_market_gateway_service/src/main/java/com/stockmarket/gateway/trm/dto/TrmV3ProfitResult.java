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
 *         TEJA RECOVERY MODEL (TRM) V3 PROFIT SCENARIO RESULT DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Result payload for TRM V3 Path A (Profit Scenario with Zero-Cost Runner Retention).
 * Enforces the <i>Current Price Law</i>:
 * <pre>
 *   Q_{retain} = RetainedProfitValue / CurrentPrice   [DIVIDED BY CURRENT PRICE, NOT ENTRY PRICE]
 *   Q_{sell} = OriginalQuantity - Q_{retain}
 *   CashRealized = Q_{sell} * CurrentPrice
 * </pre>
 * Ensures 100% of initial principal is extracted into liquid cash, while the retained shares carry
 * an effective post-reconciliation cost basis of $0.00.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok annotations.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Excludes null values.</li>
 *   <li><b>@JsonProperty</b>: Enforces consistent JSON keys.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Triggered upon execution of a profit target limit order. Generates the exact split between market
 * liquidation quantity and vault-retained runner shares.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmV3ProfitResult implements Serializable {

    /**
     * Original entry purchase price per share (P_0).
     */
    @JsonProperty("entryPrice")
    private BigDecimal entryPrice;

    /**
     * Execution price at the profit target level: EntryPrice * (1 + ProfitPercent).
     */
    @JsonProperty("currentPrice")
    private BigDecimal currentPrice;

    /**
     * Initial position volume purchased at entry.
     */
    @JsonProperty("originalQuantity")
    private BigDecimal originalQuantity;

    /**
     * Total seed cash invested: EntryPrice * OriginalQuantity.
     */
    @JsonProperty("initialInvestment")
    private BigDecimal initialInvestment;

    /**
     * Percentage gain achieved at target fill (e.g., 0.15 for 15%).
     */
    @JsonProperty("profitPercent")
    private BigDecimal profitPercent;

    /**
     * Gross dollar profit realized on the trade: InitialInvestment * ProfitPercent.
     */
    @JsonProperty("profitAmount")
    private BigDecimal profitAmount;

    /**
     * Fraction of net profit converted into permanent runner equity (e.g., 0.15).
     */
    @JsonProperty("retentionFactor")
    private BigDecimal retentionFactor;

    /**
     * Monetary value of profit allocated to acquiring/retaining runner shares.
     */
    @JsonProperty("retainedProfitValue")
    private BigDecimal retainedProfitValue;

    /**
     * Precise number of runner shares retained under Current Price Law: RetainedProfitValue / CurrentPrice.
     */
    @JsonProperty("retainedQuantity")
    private BigDecimal retainedQuantity;

    /**
     * Number of shares sold to the market: OriginalQuantity - RetainedQuantity.
     */
    @JsonProperty("quantityToSell")
    private BigDecimal quantityToSell;

    /**
     * Total gross cash returned from selling: QuantityToSell * CurrentPrice.
     */
    @JsonProperty("cashRealized")
    private BigDecimal cashRealized;

    /**
     * Portion of cash realized representing 100% of the original principal investment.
     */
    @JsonProperty("initialCapitalReturned")
    private BigDecimal initialCapitalReturned;

    /**
     * Net liquid cash profit harvested: ProfitAmount * (1 - RetentionFactor).
     */
    @JsonProperty("netCashProfit")
    private BigDecimal netCashProfit;

    /**
     * Current market value of the permanently retained runner shares: RetainedQuantity * CurrentPrice.
     */
    @JsonProperty("runnerEquityValue")
    private BigDecimal runnerEquityValue;

    /**
     * Effective cost basis of the runner position ($0.00 since principal has been fully extracted).
     */
    @JsonProperty("effectiveCostBasis")
    private BigDecimal effectiveCostBasis;

    /**
     * Mathematical reconciliation check: CashRealized + RunnerEquityValue == OriginalQuantity * CurrentPrice.
     */
    @JsonProperty("reconciliationReconciled")
    private boolean reconciliationReconciled;

    /**
     * Action mandate instructing the broker/trader of execution steps.
     */
    @JsonProperty("actionMandate")
    private String actionMandate;
}
