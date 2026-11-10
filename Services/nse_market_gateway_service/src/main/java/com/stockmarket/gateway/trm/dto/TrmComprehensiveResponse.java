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
 *         TEJA RECOVERY MODEL (TRM) COMPREHENSIVE MULTI-ENGINE EVALUATION DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Top-level composite Data Transfer Object unifies Subsystem A (Statistical Confluence & OLS Linear Regression Channel)
 * with Subsystem B (concurrent execution outputs from all 4 TRM quantitative engines: Classic, V1, V2, and V3).
 * Formulates a side-by-side comparative summary and an actionable executive directive.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok boilerplate.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Excludes unpopulated fields.</li>
 *   <li><b>@JsonProperty</b>: Standardizes JSON serialization keys.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * The single unified payload requested by high-level portfolio managers, automated trading systems, and
 * GenAI decision agents before committing capital to a live stock setup.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmComprehensiveResponse implements Serializable {

    /**
     * NSE equity ticker symbol evaluated (e.g., "RELIANCE").
     */
    @JsonProperty("symbol")
    private String symbol;

    /**
     * ISO-8601 UTC timestamp of the evaluation run.
     */
    @JsonProperty("timestamp")
    private String timestamp;

    /**
     * Technical regime metrics including 100-period OLS regression channel, VWAP distance, and setup score.
     */
    @JsonProperty("technicalConfluence")
    private TrmTechnicalConfluence technicalConfluence;

    /**
     * Model 1 Output: TRM Classic recursive cumulative capital progression.
     */
    @JsonProperty("trmClassic")
    private TrmClassicResponse trmClassic;

    /**
     * Model 2 Output: TRM V1 percentage-triggered stock equity progression.
     */
    @JsonProperty("trmV1")
    private TrmV1Response trmV1;

    /**
     * Model 3 Output: TRM V2 actual-portfolio-loss adaptive progression (~99.68% capital savings).
     */
    @JsonProperty("trmV2")
    private TrmV2Response trmV2;

    /**
     * Model 4 Output: TRM V3 asymmetric dual-path engine (Zero-cost runner retention + recovery surcharge).
     */
    @JsonProperty("trmV3")
    private TrmV3Response trmV3;

    /**
     * Quantitative comparison summary contrasting capital requirements and efficiency across all 4 models.
     */
    @JsonProperty("comparisonSummary")
    private String comparisonSummary;

    /**
     * Actionable trading mandate (e.g., "HIGH CONVICTION VALUE ENTRY", "OVERBOUGHT CEILING: STRICTLY NO ENTRY").
     */
    @JsonProperty("executiveDirective")
    private String executiveDirective;
}
