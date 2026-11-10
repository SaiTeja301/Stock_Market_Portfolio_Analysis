package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import java.io.Serializable;

/**
 * ====================================================================================================
 *                 AI MODEL CATALOG DTO: MODEL METADATA RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object representing the technical capabilities, hardware specifications,
 * VRAM footprints, and operational ratings of LLMs available in the Ollama / Spring AI / LangChain4j ecosystem.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonProperty</b>: Guarantees standardized JSON serialization across REST endpoints.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Presented in API catalogs (e.g., {@code GET /api/v1/ai/models}) allowing client applications and trading
 * desks to inspect available models (DeepSeek-R1 1.5B, Qwen 2.5 3B, Gemma 2 2B, Llama 3.2 3B) and dynamically
 * route tasks to the model best suited for their latency and reasoning needs.
 *
 * @param modelId        Unique identifier string of the model (e.g. "qwen2.5:3b", "deepseek-r1:1.5b").
 * @param displayName    Human-readable name of the model.
 * @param provider       Underlying inference engine provider (e.g., "Ollama (Local)", "Spring AI").
 * @param category       Model architectural focus (e.g., "Reasoning & Mathematical Verification", "General Speed").
 * @param description    Overview of the model's strengths and neural architecture.
 * @param recommendedFor Specific trading or coding use cases where this model excels.
 * @param vramUsage      Approximate GPU VRAM required during inference execution.
 * @param speedRating    Throughput rating (e.g., "Fast (~45 tok/s)", "Moderate").
 */
public record ModelMetadataRecord(
        @JsonProperty("modelId") String modelId,
        @JsonProperty("displayName") String displayName,
        @JsonProperty("provider") String provider,
        @JsonProperty("category") String category,
        @JsonProperty("description") String description,
        @JsonProperty("recommendedFor") String recommendedFor,
        @JsonProperty("vramUsage") String vramUsage,
        @JsonProperty("speedRating") String speedRating
) implements Serializable {}
