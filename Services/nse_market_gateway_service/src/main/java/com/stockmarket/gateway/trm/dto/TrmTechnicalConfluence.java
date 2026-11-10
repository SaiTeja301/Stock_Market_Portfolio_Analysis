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
 *         TRM SUBSYSTEM A: TECHNICAL CONFLUENCE & REGRESSION CHANNEL DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Data Transfer Object encapsulating Subsystem A (Trend-Regression-Momentum Confluence).
 * Combines live Level-1 exchange quotes with a 100-period Ordinary Least Squares (OLS) Linear Regression
 * Channel (\u00b12.0\u03c3), an 8-dimension 100-point setup scorecard, and a pre-flight execution checklist.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok annotations.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Excludes unpopulated fields.</li>
 *   <li><b>@JsonProperty</b>: Standardizes JSON serialization keys.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Consumed by automated trade filtering rules to confirm whether an equity is in a qualified demand zone
 * (-2.0\u03c3 Lower Band) or overbought ceiling (+2.0\u03c3 Upper Band) before triggering TRM recovery tranches.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmTechnicalConfluence implements Serializable {

    /**
     * NSE equity ticker symbol (e.g., "INFY").
     */
    @JsonProperty("symbol")
    private String symbol;

    /**
     * Last Traded Price (LTP) on the exchange matching engine.
     */
    @JsonProperty("lastTradedPrice")
    private BigDecimal lastTradedPrice;

    /**
     * Net session change from previous close.
     */
    @JsonProperty("change")
    private BigDecimal change;

    /**
     * Percentage session change from previous close.
     */
    @JsonProperty("pChange")
    private BigDecimal pChange;

    /**
     * Pre-market open print price.
     */
    @JsonProperty("openPrice")
    private BigDecimal openPrice;

    /**
     * Prior regular session closing price.
     */
    @JsonProperty("closePrice")
    private BigDecimal closePrice;

    /**
     * Intraday Volume Weighted Average Price.
     */
    @JsonProperty("vwap")
    private BigDecimal vwap;

    /**
     * Daily circuit limit lower price floor.
     */
    @JsonProperty("lowerCircuit")
    private BigDecimal lowerCircuit;

    /**
     * Daily circuit limit upper price ceiling.
     */
    @JsonProperty("upperCircuit")
    private BigDecimal upperCircuit;

    /**
     * 52-week peak price.
     */
    @JsonProperty("week52High")
    private BigDecimal week52High;

    /**
     * 52-week trough price.
     */
    @JsonProperty("week52Low")
    private BigDecimal week52Low;

    /**
     * OLS Linear regression slope (\u03b2): rate of price change per bar.
     */
    @JsonProperty("regressionSlope")
    private BigDecimal regressionSlope;

    /**
     * OLS Linear regression vertical intercept (\u03b1).
     */
    @JsonProperty("regressionIntercept")
    private BigDecimal regressionIntercept;

    /**
     * Fitted OLS regression trendline mean price at the current bar: \u03b1 + \u03b2 * n.
     */
    @JsonProperty("regressionMean")
    private BigDecimal regressionMean;

    /**
     * Standard error of estimate (\u03c3) around the OLS regression trendline.
     */
    @JsonProperty("standardError")
    private BigDecimal standardError;

    /**
     * Upper Bollinger/regression envelope boundary (+2.0\u03c3 overbought ceiling).
     */
    @JsonProperty("upperBand2Sigma")
    private BigDecimal upperBand2Sigma;

    /**
     * Lower Bollinger/regression envelope boundary (-2.0\u03c3 demand entry floor).
     */
    @JsonProperty("lowerBand2Sigma")
    private BigDecimal lowerBand2Sigma;

    /**
     * Categorical channel positioning (e.g. "LOWER_BAND_DEMAND_ZONE (-2.0\u03c3 High Probability Entry)").
     */
    @JsonProperty("channelPosition")
    private String channelPosition;

    /**
     * 100-point setup score derived from the 8-dimension quantitative rubric.
     */
    @JsonProperty("setupScore")
    private int setupScore;

    /**
     * Categorical classification: "STRONG", "GOOD", "WATCHLIST", or "AVOID".
     */
    @JsonProperty("setupClassification")
    private String setupClassification;

    /**
     * 8-point pre-flight validation checklist ensuring risk compliance before execution.
     */
    @JsonProperty("preFlightChecklist")
    private List<String> preFlightChecklist;
}
