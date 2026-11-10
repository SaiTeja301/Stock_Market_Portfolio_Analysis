package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: TOP GAINER / LOSER RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object capturing performance and turnover metrics for an individual
 * top-gaining or top-losing stock in the market breadth index.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Jackson annotation protecting against unknown fields.</li>
 *   <li><b>@JsonProperty</b>: Standardized JSON property mapping.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Used by momentum breakout strategies to identify the strongest intraday runners, and by mean-reversion
 * strategies to locate oversold bounce candidates trading near their daily circuit bounds.
 *
 * @param symbol            Official NSE equity ticker symbol.
 * @param series            Trading category series (typically "EQ").
 * @param lastPrice         Last Traded Price.
 * @param change            Absolute net price movement from previous close.
 * @param pChange           Percentage price movement from previous close.
 * @param totalTradedVolume Total cumulative shares exchanged during the session.
 * @param totalTradedValue  Total cumulative turnover value in INR Crores.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record TopGainerLoserRecord(
        @JsonProperty("symbol") String symbol,
        @JsonProperty("series") String series,
        @JsonProperty("lastPrice") double lastPrice,
        @JsonProperty("change") double change,
        @JsonProperty("pChange") double pChange,
        @JsonProperty("totalTradedVolume") long totalTradedVolume,
        @JsonProperty("totalTradedValue") double totalTradedValue
) implements Serializable {}
