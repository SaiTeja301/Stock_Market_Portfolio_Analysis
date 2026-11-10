package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: SYMBOL AUTOCOMPLETE SEARCH RESULT RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object capturing an individual autocomplete match result when querying
 * equities or indices on the National Stock Exchange.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Silently drops extraneous search metadata.</li>
 *   <li><b>@JsonProperty</b>: Standardized JSON property binding.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Powers instantaneous client-side UI autocomplete dropdowns as traders type ticker symbols or company
 * names into search bars, enabling quick selection before dispatching TRM or AI advisory queries.
 *
 * @param symbol      Official NSE equity ticker symbol (e.g., "INFY").
 * @param companyName Full legal name of the company (e.g., "Infosys Limited").
 * @param series      Trading category (e.g., "EQ", "BE").
 * @param segment     Market segment identifier (e.g., "Equity", "Index").
 * @param lastPrice   Latest traded price at the time of indexation.
 * @param url         Exchange URL slug for corporate profiling.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record SearchResultRecord(
        @JsonProperty("symbol") String symbol,
        @JsonProperty("companyName") String companyName,
        @JsonProperty("series") String series,
        @JsonProperty("segment") String segment,
        @JsonProperty("lastPrice") double lastPrice,
        @JsonProperty("url") String url
) implements Serializable {}
