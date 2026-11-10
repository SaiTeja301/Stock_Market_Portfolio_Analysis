package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;
import java.util.List;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: ORDER BOOK (DEPTH OF MARKET) RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object representing the Level-2 Depth of Market (DOM) snapshot.
 * Aggregates the top 5 bid and ask limit price tiers alongside the total cumulative buy and sell
 * order quantities queued in the exchange matching engine.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Jackson annotation preventing deserialization
 *       failures when NSE returns supplemental market depth attributes.</li>
 *   <li><b>@JsonProperty</b>: Binds JSON attributes directly to record components.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Deployed in order flow imbalance analysis and liquidity sweeps. If {@code totalBuyQuantity} vastly
 * exceeds {@code totalSellQuantity} (e.g. 3:1 ratio), algorithmic market makers identify aggressive
 * buying pressure, assisting the TRM engine in validating high-probability demand zone entries.
 *
 * @param levels            Top-5 depth tiers containing simultaneous bid and ask price/quantity pairs.
 * @param totalBuyQuantity  Aggregate volume of all active buy limit orders visible in the book.
 * @param totalSellQuantity Aggregate volume of all active sell limit orders visible in the book.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record OrderBookRecord(
        @JsonProperty("levels") List<OrderBookLevelRecord> levels,
        @JsonProperty("totalBuyQuantity") long totalBuyQuantity,
        @JsonProperty("totalSellQuantity") long totalSellQuantity
) implements Serializable {}
