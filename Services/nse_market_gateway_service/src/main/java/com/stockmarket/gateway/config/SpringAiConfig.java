package com.stockmarket.gateway.config;

import org.springframework.ai.chat.client.ChatClient;
import org.springframework.ai.chat.model.ChatModel;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * ====================================================================================================
 *                               SPRING AI ENTERPRISE CONFIGURATION
 * ====================================================================================================
 * 
 * Architectural Overview:
 * Spring AI provides an idiomatic, enterprise-grade abstraction layer for interacting with Artificial
 * Intelligence models (OpenAI, Ollama, Anthropic, Bedrock, etc.) natively in Spring Boot 3.
 * 
 * Key Spring AI Concepts Configured Here:
 * 
 * 1. ChatModel (org.springframework.ai.chat.model.ChatModel):
 *    - Role: The low-level provider SPI interface. In this microservice, Spring Boot autoconfigures
 *      `OllamaChatModel` from the starter `spring-ai-ollama-spring-boot-starter`.
 *    - Mechanics: Handles HTTP serialization to Ollama's `/api/chat` endpoint, parameterizing temperature,
 *      top-p, and model names (`qwen2.5:3b`, `deepseek-r1:1.5b`, `llama3.2:3b`, `gemma2:2b`).
 * 
 * 2. ChatClient (org.springframework.ai.chat.client.ChatClient):
 *    - Role: Spring AI's primary high-level fluent API (modeled after `WebClient`, `RestClient`, and `JdbcClient`).
 *    - Mechanics: Created via `ChatClient.builder(chatModel)`. Enables declarative chaining:
 *      {@code chatClient.prompt().system(...).user(...).tools(...).call().content()}.
 *    - Default System Instructions: Injected via `.defaultSystem(...)` to govern every prompt across the service.
 * 
 * Architectural Comparison: Spring AI vs LangChain4j:
 * - Spring AI uses a fluent builder API (`ChatClient`) with functional tool callbacks (`Function<Req, Resp>`).
 * - LangChain4j uses declarative annotated interfaces (`@AiService`) with `@Tool` method annotations.
 * - This microservice showcases both enterprise paradigms running side-by-side on the same JVM.
 */
@Configuration
public class SpringAiConfig {

    /**
     * Instantiates and registers the primary {@link ChatClient} bean in the Spring container.
     * 
     * Annotation: @Bean
     * Role: Exposes the ChatClient singleton bean for injection into Spring services (e.g. {@link com.stockmarket.gateway.service.SpringAiAdvisoryService}).
     * 
     * Configuration:
     * Enforces the institutional TRM Persona, 100-period regression bounds, and the 48-hour binary event embargo.
     * 
     * @param chatModel The Spring-managed ChatModel auto-configured by Spring AI Ollama starter.
     * @return Configured ChatClient singleton instance.
     */
    @Bean
    public ChatClient chatClient(ChatModel chatModel) {
        return ChatClient.builder(chatModel)
                .defaultSystem("""
                        You are a Senior Indian Equity Market Analyst, Quantitative Strategist,
                        and Teja Recovery Model (TRM) Algorithm Expert.
                        Your analysis strictly follows the TRM Strategy:
                        1. Multi-timeframe trend alignment (Daily + 15M/Hourly)
                        2. 100-period OLS Linear Regression Channel positioning (±2.0σ demand/exhaustion bands)
                        3. Momentum confluence (RSI 40-50 pullback bounce, MACD zero line)
                        4. Binary event risk embargo: 48-hour trading freeze before scheduled earnings or board meetings
                        5. Strict data provenance: Tag all output with [OBSERVED], [CALCULATED], [ASSUMED].
                        Always evaluate live NSE quote and corporate announcement data.
                        """)
                .build();
    }
}
