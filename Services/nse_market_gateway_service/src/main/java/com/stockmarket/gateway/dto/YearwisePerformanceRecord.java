package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: MULTI-HORIZON HISTORICAL PERFORMANCE RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object encapsulating historical returns across multiple standard lookback
 * windows (1 day to 5 years), mapped alongside benchmark index returns for alpha/beta evaluation.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Silently drops extraneous return metrics.</li>
 *   <li><b>@JsonProperty</b>: Guarantees standardized JSON serialization.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Quantitative equity screeners analyze medium-term (3-month / 6-month) vs long-term (1-year / 3-year)
 * performance to classify equities into Momentum Leaders, Mean Reversion Laggards, or Long-Term Compounders.
 *
 * @param yesterdayChngPer    1-day return percentage.
 * @param oneWeekChngPer      1-week trailing return percentage.
 * @param oneMonthChngPer     1-month trailing return percentage.
 * @param threeMonthChngPer   3-month trailing return percentage.
 * @param sixMonthChngPer     6-month trailing return percentage.
 * @param oneYearChngPer      1-year trailing return percentage.
 * @param twoYearChngPer      2-year trailing compound return percentage.
 * @param threeYearChngPer    3-year trailing compound return percentage.
 * @param fiveYearChngPer     5-year trailing compound return percentage.
 * @param indexName           Benchmark index symbol for comparison (e.g., "NIFTY 50").
 * @param indexOneYearChngPer 1-year trailing performance of the benchmark index.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record YearwisePerformanceRecord(
        @JsonProperty("yesterdayChngPer") double yesterdayChngPer,
        @JsonProperty("oneWeekChngPer") double oneWeekChngPer,
        @JsonProperty("oneMonthChngPer") double oneMonthChngPer,
        @JsonProperty("threeMonthChngPer") double threeMonthChngPer,
        @JsonProperty("sixMonthChngPer") double sixMonthChngPer,
        @JsonProperty("oneYearChngPer") double oneYearChngPer,
        @JsonProperty("twoYearChngPer") double twoYearChngPer,
        @JsonProperty("threeYearChngPer") double threeYearChngPer,
        @JsonProperty("fiveYearChngPer") double fiveYearChngPer,
        @JsonProperty("indexName") String indexName,
        @JsonProperty("indexOneYearChngPer") double indexOneYearChngPer
) implements Serializable {}
