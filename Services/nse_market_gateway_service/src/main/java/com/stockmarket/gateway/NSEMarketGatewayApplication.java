package com.stockmarket.gateway;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.cache.annotation.EnableCaching;
import org.springframework.scheduling.annotation.EnableScheduling;

/**
 * ====================================================================================================
 *                 NSE MARKET GATEWAY MICROSERVICE: ENTRY POINT
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Main bootstrap and entry point for the NSE Market Gateway Microservice.
 * Co-locates three enterprise architectural engines:
 * 1. Real-Time NSE Market Data Gateway: Ingests, normalizes, and caches Level-2 order book depth,
 *    intraday charts, corporate filings, and benchmark index membership from the National Stock Exchange of India.
 * 2. Teja Recovery Model (TRM) Quantitative Algorithmic Suite: Deterministic mathematical implementations
 *    of TRM Classic (geometric tripling), TRM V1 (percentage-triggered), TRM V2 (actual-portfolio-loss adaptive),
 *    and TRM V3 (asymmetric profit runner vaulting + recovery surcharge), combined with 100-period OLS linear regression channels.
 * 3. Dual-Engine Generative AI Advisory: Spring AI dynamic multi-model routing (Qwen 2.5 3B, DeepSeek-R1 1.5B,
 *    Gemma 2 2B, Llama 3.2 3B) alongside LangChain4j autonomous agents with declarative {@code @Tool} dispatch,
 *    stateful session memory ({@code @MemoryId}), and typed structured extraction.
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@SpringBootApplication</b>: Meta-annotation combining {@code @Configuration}, {@code @EnableAutoConfiguration},
 *       and {@code @ComponentScan}. Scans for Spring components under the root package {@code com.stockmarket.gateway}.</li>
 *   <li><b>@EnableScheduling</b>: Enables Spring's asynchronous scheduled task execution infrastructure,
 *       activating {@link com.stockmarket.gateway.session.NSESessionManager#scheduledWarmup()} for periodic Akamai cookie rotation.</li>
 *   <li><b>@EnableCaching</b>: Activates Spring's declarative annotation-driven cache management infrastructure,
 *       binding {@code @Cacheable} methods to the Caffeine in-memory cache manager configured in {@link com.stockmarket.gateway.config.CacheConfig}.</li>
 * </ul>
 */
@SpringBootApplication
@EnableScheduling
@EnableCaching
public class NSEMarketGatewayApplication {

    /**
     * Standard JVM entry point launching the Spring Boot embedded Tomcat web server on port 8081.
     *
     * @param args Command-line arguments passed during application startup.
     */
    public static void main(String[] args) {
        SpringApplication.run(NSEMarketGatewayApplication.class, args);
    }
}
