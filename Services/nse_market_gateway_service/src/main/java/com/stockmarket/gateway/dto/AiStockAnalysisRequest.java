package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import jakarta.validation.constraints.NotBlank;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 AI ADVISORY DTO: STOCK ANALYSIS REQUEST RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object defining the client request payload for AI-powered equity research,
 * TRM quantitative advisory synthesis, and dynamic multi-model reasoning.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@NotBlank</b>: Jakarta validation constraint guaranteeing that a non-null, non-empty ticker
 *       symbol is supplied, triggering HTTP 400 Bad Request if omitted.</li>
 *   <li><b>@JsonProperty</b>: Standardized JSON property mapping.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Dispatched by traders or portfolio management systems to request an autonomous end-to-end investment
 * analysis on an equity symbol (e.g. "RELIANCE") using a specified model (e.g., "deepseek-r1:1.5b")
 * and analysis focus (e.g., "comprehensive", "trm", "earnings").
 *
 * @param symbol       Mandatory NSE equity ticker symbol to analyze (e.g. "TATAMOTORS").
 * @param analysisType Category of analysis requested (e.g. "COMPREHENSIVE", "TECHNICAL", "TRM_RECOVERY").
 * @param model        Optional LLM model override; if omitted, intelligent rule-based routing selects the model.
 * @param userContext  Optional custom context supplied by the trader (e.g. "Holdings: 500 shares at 950 INR").
 */
public record AiStockAnalysisRequest(
        @NotBlank(message = "Stock symbol is required")
        @JsonProperty("symbol") String symbol,
        @JsonProperty("analysisType") String analysisType,
        @JsonProperty("model") String model,
        @JsonProperty("userContext") String userContext
) implements Serializable {}
