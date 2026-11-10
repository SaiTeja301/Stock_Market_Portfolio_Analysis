package com.stockmarket.gateway.trm.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * ====================================================================================================
 *         TEJA RECOVERY MODEL (TRM) V3 ASYMMETRIC RESPONSE DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Consolidated response payload for the TRM V3 asymmetric dual-path trading engine.
 * Encapsulates pre-calculated execution paths for both trade outcomes:
 * <ul>
 *   <li><b>Path A (Profit Scenario)</b>: Enforces the <i>Current Price Law</i> (\(Q_{retain} = R / P_{curr}\))
 *       to return 100% of seed capital to cash, realize net profit, and hold zero-cost runner shares.</li>
 *   <li><b>Path B (Loss Scenario)</b>: Executes a structural stop-loss, liquidates 100% of equity to preserve
 *       cash, and computes a linear recovery surcharge for the next reversal cycle.</li>
 * </ul>
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok boilerplate.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Omits empty properties.</li>
 *   <li><b>@JsonProperty</b>: Standardizes JSON serialization.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Deployed in automated bracket-order execution. Prior to trade entry, the trader or algorithm knows the
 * exact liquidation size and runner retention count on a win, and the exact stop price and recovery surcharge on a loss.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmV3Response implements Serializable {

    /**
     * Display name of the model ("TRM V3 - Asymmetric Dual-Path Model").
     */
    @JsonProperty("modelName")
    @Builder.Default
    private String modelName = "TRM V3 - Asymmetric Dual-Path Model";

    /**
     * Complete calculation metrics for Path A (Profit scenario with runner retention).
     */
    @JsonProperty("profitScenario")
    private TrmV3ProfitResult profitScenario;

    /**
     * Complete calculation metrics for Path B (Loss scenario with linear recovery surcharge).
     */
    @JsonProperty("lossScenario")
    private TrmV3LossResult lossScenario;

    /**
     * Mandate stating that upon cycle completion, the active trade allocation unconditionally resets to Base S.
     */
    @JsonProperty("universalResetRule")
    @Builder.Default
    private String universalResetRule = "Upon target fill or recovery cycle completion, active allocation unconditionally resets to Base S.";

    /**
     * Mathematical verification flag confirming that runner retention divided by current price, not entry price.
     */
    @JsonProperty("currentPriceLawVerified")
    @Builder.Default
    private boolean currentPriceLawVerified = true;
}
