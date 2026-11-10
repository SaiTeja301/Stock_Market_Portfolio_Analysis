package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;
import java.util.List;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: HISTORICAL & INTRADAY SYMBOL CHART RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object encapsulating time-series price points for an NSE equity symbol.
 * Powers technical indicator calculations, chart plotting, and OLS regression channel modeling.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Jackson annotation ignoring supplemental metadata.</li>
 *   <li><b>@JsonProperty</b>: Standardized JSON property binding.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Feeds the 100-period Ordinary Least Squares (OLS) regression engine in TRM Subsystem A to calculate
 * regression mean, standard error bands (\u00b12.0\u03c3), and current channel position for breakout/mean-reversion setups.
 *
 * @param identifier Official ticker identifier string (e.g., "INFYEQN").
 * @param name       Full corporate name of the issuing company.
 * @param closePrice Session reference closing price.
 * @param points     Chronological list of {@link ChartPointRecord} entries.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record SymbolChartRecord(
        @JsonProperty("identifier") String identifier,
        @JsonProperty("name") String name,
        @JsonProperty("closePrice") double closePrice,
        @JsonProperty("points") List<ChartPointRecord> points
) implements Serializable {}
