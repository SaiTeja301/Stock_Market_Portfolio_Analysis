package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import dev.langchain4j.model.output.structured.Description;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 LANGCHAIN4J STRUCTURED OUTPUT: TRM ADVISORY RECORD
 * ====================================================================================================
 * 
 * Architectural Role:
 * Represents the strongly-typed structured output schema extracted directly by LangChain4j's
 * {@code @AiService} engine.
 * 
 * Key LangChain4j Concepts Demonstrated:
 * 1. Structured JSON Schema Extraction (Type-Safe LLM Responses):
 *    In standard LLM interactions, the model returns arbitrary free-form Markdown text. Extracting
 *    trading levels like Stop-Loss or Target prices usually requires fragile regex pattern matching.
 *    LangChain4j solves this natively: by returning this Record directly from an {@code @AiService} method,
 *    LangChain4j automatically inspects its reflection metadata, constructs a strict JSON Schema,
 *    prompts the LLM (e.g. Ollama Qwen/DeepSeek/Llama) to output JSON conforming to that schema, and
 *    deserializes the response directly into this strongly-typed Java Record.
 * 
 * 2. dev.langchain4j.model.output.structured.Description:
 *    Instructs the LLM via schema metadata what each individual property represents, eliminating ambiguity.
 * 
 * Real-Time Production Scenario:
 * Algorithmic order routers and risk engines cannot execute trade orders from unstructured English paragraphs.
 * They require exact numeric floats/doubles (e.g. stopLoss = 2860.00, recommendation = "STRONG_BUY").
 * This Record bridges the non-deterministic LLM world with deterministic trading microservices.
 */
public record TrmStructuredAdvisory(

        @Description("The NSE ticker symbol, e.g. RELIANCE, TCS, INFY")
        @JsonProperty("symbol")
        String symbol,

        @Description("Last Traded Price (LTP) observed from live market data")
        @JsonProperty("lastTradedPrice")
        double lastTradedPrice,

        @Description("Calculated TRM structural stop-loss price floor (e.g. 4% below entry or below channel)")
        @JsonProperty("trmStopLoss")
        double trmStopLoss,

        @Description("Conservative Target 1 price level (e.g. +5% or Regression Channel Mean)")
        @JsonProperty("trmTarget1")
        double trmTarget1,

        @Description("Base Case Target 2 price level (e.g. +10% or Upper Regression Band +2.0σ)")
        @JsonProperty("trmTarget2")
        double trmTarget2,

        @Description("Objective quantitative setup quality score from 0 to 100 based on the 8-dimension rubric")
        @JsonProperty("setupScore")
        int setupScore,

        @Description("Executive action recommendation: STRONG_BUY, TRM_RE_ENTRY, HOLD, PARTIAL_EXIT, or AVOID")
        @JsonProperty("recommendation")
        String recommendation,

        @Description("Crisp quantitative rationale synthesizing trend, channel position, and volume")
        @JsonProperty("keyRationale")
        String keyRationale,

        @Description("True if no board meetings or financial results are scheduled within 48 hours; false otherwise")
        @JsonProperty("binaryEventEmbargoClear")
        boolean binaryEventEmbargoClear

) implements Serializable {}
