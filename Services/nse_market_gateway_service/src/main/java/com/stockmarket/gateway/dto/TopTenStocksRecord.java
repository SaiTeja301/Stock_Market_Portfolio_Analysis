package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: TOP TEN STOCKS MARKET BREADTH RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object capturing exchange-wide market movers, liquidity leaders, and unusual
 * volume anomalies across the National Stock Exchange (NSE) equity segment.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Silently drops extraneous market statistics.</li>
 *   <li><b>@JsonProperty</b>: Guarantees standardized JSON mapping.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Algorithmic scanners and AI agents evaluate top gainers, losers, and volume spurts during the market
 * session to detect sector rotation, institutional accumulation, and sudden volatility spikes.
 *
 * @param timestamp        Broadcast timestamp of the market breadth report.
 * @param topGainers       List of top advancing stocks by percentage price appreciation.
 * @param topLosers        List of top declining stocks by percentage price depreciation.
 * @param mostActiveValue  List of stocks with highest turnover (traded value in Crores).
 * @param mostActiveVolume List of stocks with highest traded share volume.
 * @param volumeSpurts     Stocks experiencing sudden intraday volume surges relative to trailing averages.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record TopTenStocksRecord(
        @JsonProperty("timestamp") String timestamp,
        @JsonProperty("topGainers") List<TopGainerLoserRecord> topGainers,
        @JsonProperty("topLosers") List<TopGainerLoserRecord> topLosers,
        @JsonProperty("mostActiveValue") List<Map<String, Object>> mostActiveValue,
        @JsonProperty("mostActiveVolume") List<Map<String, Object>> mostActiveVolume,
        @JsonProperty("volumeSpurts") List<Map<String, Object>> volumeSpurts
) implements Serializable {}
