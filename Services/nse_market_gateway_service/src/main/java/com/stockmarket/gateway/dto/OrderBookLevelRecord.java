package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: ORDER BOOK LEVEL DEPTH TIER RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Represents an individual tier (1 to 5) within the Depth of Market (DOM) order book hierarchy.
 * Captures simultaneous bid (buy) and ask (sell) limit prices alongside queued quantities.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Protects against unexpected attributes.</li>
 *   <li><b>@JsonProperty</b>: Guarantees standardized JSON serialization.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Level 1 represents the Best Bid and Offer (BBO). The spread between {@code buyPrice} and {@code sellPrice}
 * determines the immediate slippage cost for market orders. TRM execution algorithms verify Level 1-3 depth
 * before submitting tranche orders to ensure adequate liquidity without price impact.
 *
 * @param level        Depth level ordinal (1 = best bid/ask, 5 = furthest depth tier).
 * @param buyPrice     Bid limit price queued at this level.
 * @param buyQuantity  Total share quantity queued at this bid price.
 * @param sellPrice    Ask/offer limit price queued at this level.
 * @param sellQuantity Total share quantity queued at this ask price.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record OrderBookLevelRecord(
        @JsonProperty("level") int level,
        @JsonProperty("buyPrice") double buyPrice,
        @JsonProperty("buyQuantity") long buyQuantity,
        @JsonProperty("sellPrice") double sellPrice,
        @JsonProperty("sellQuantity") long sellQuantity
) implements Serializable {}
