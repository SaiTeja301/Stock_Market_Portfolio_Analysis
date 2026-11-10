package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: INTRADAY TIME-SERIES CHART POINT RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Represents a discrete price snapshot point in a financial time-series chart.
 * Captures epoch millisecond timestamp, executed price, trading session state, and net changes.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Silently skips auxiliary charting fields.</li>
 *   <li><b>@JsonProperty</b>: Guarantees precise JSON serialization.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Used to reconstruct intraday OHLCV bars and feed mathematical regression models. Traders observe
 * the delta and tradeState (e.g. "REGULAR", "AUCTION") to confirm volume-backed price trends.
 *
 * @param timestampMs Epoch millisecond timestamp of the price print.
 * @param price       Execution price at this timestamp.
 * @param tradeState  Trading session status (e.g. "REGULAR", "CLOSED", "HALTED").
 * @param change      Net price variation from the prior reference close.
 * @param pChange     Percentage price variation from the prior reference close.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record ChartPointRecord(
        @JsonProperty("timestampMs") long timestampMs,
        @JsonProperty("price") double price,
        @JsonProperty("tradeState") String tradeState,
        @JsonProperty("change") double change,
        @JsonProperty("pChange") double pChange
) implements Serializable {}
