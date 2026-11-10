package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: REAL-TIME SYMBOL QUOTE RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object (Record) encapsulating real-time Level-1 equity quote data from the
 * National Stock Exchange (NSE). Serves as the single source of truth for stock valuation across the
 * microservice, feeding the TRM quantitative math engines, caching layers, and GenAI advisory pipelines.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Jackson annotation instructing deserializers
 *       to ignore unknown or undocumented fields returned by upstream NSE endpoints (e.g. meta tags,
 *       internal security identifiers), preventing JSON deserialization breaks during upstream schema drift.</li>
 *   <li><b>@JsonProperty("fieldName")</b>: Maps JSON key names explicitly to Java record components,
 *       guaranteeing cross-platform serialization compatibility.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Consumed by algorithmic execution bots and AI agents to check the Last Traded Price (LTP), determine if
 * the stock is trading near daily circuit limits (lowerCircuit/upperCircuit), assess intraday volume
 * weight (VWAP), and inspect top-5 order book depth before firing limit orders.
 *
 * @param symbol        Official NSE equity ticker symbol (e.g., "RELIANCE", "TCS").
 * @param series        Trading series category (e.g., "EQ" for standard equity settlement, "BE" for trade-to-trade).
 * @param lastPrice     Last Traded Price (LTP) executed on the exchange matching engine.
 * @param change        Absolute net price movement relative to previous session close.
 * @param pChange       Percentage price movement relative to previous close.
 * @param previousClose Official closing price from the preceding regular market session.
 * @param openPrice     Session opening print established during the 09:00 - 09:08 IST pre-market auction.
 * @param closePrice    Final calculated session closing price (weighted average of final 30 minutes).
 * @param vwap          Volume Weighted Average Price across the entire intraday session.
 * @param lowerCircuit  Exchange-mandated daily lower circuit limit (automatic trading halt price floor).
 * @param upperCircuit  Exchange-mandated daily upper circuit limit (automatic trading halt price ceiling).
 * @param week52High    Trailing 52-week peak traded price.
 * @param week52Low     Trailing 52-week trough traded price.
 * @param orderBook     Nested {@link OrderBookRecord} containing top 5 bids and asks.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record SymbolQuoteRecord(
        @JsonProperty("symbol") String symbol,
        @JsonProperty("series") String series,
        @JsonProperty("lastPrice") double lastPrice,
        @JsonProperty("change") double change,
        @JsonProperty("pChange") double pChange,
        @JsonProperty("previousClose") double previousClose,
        @JsonProperty("openPrice") double openPrice,
        @JsonProperty("closePrice") double closePrice,
        @JsonProperty("vwap") double vwap,
        @JsonProperty("lowerCircuit") double lowerCircuit,
        @JsonProperty("upperCircuit") double upperCircuit,
        @JsonProperty("week52High") double week52High,
        @JsonProperty("week52Low") double week52Low,
        @JsonProperty("orderBook") OrderBookRecord orderBook
) implements Serializable {}
