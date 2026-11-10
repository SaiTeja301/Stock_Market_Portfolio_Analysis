package com.stockmarket.gateway.config;

import com.github.benmanes.caffeine.cache.Caffeine;
import org.springframework.cache.CacheManager;
import org.springframework.cache.annotation.EnableCaching;
import org.springframework.cache.caffeine.CaffeineCacheManager;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.util.concurrent.TimeUnit;

/**
 * ====================================================================================================
 *                 CAFFEINE MULTI-TIER IN-MEMORY CACHE CONFIGURATION
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Configures an ultra-high performance, low-latency in-memory caching tier powered by Ben Manes'
 * Caffeine caching library.
 *
 * Why Caffeine for Stock Market Microservices:
 * 1. Near-Zero GC Overhead: Uses Window TinyLFU (W-TinyLFU) eviction policies that outperform traditional
 *    LRU by up to 2.5x in cache hit ratio while avoiding object allocation churn.
 * 2. Sub-Microsecond Read Latency: Lock-free atomic reads allow thousands of concurrent trading threads
 *    to read quotes without blocking execution loops.
 * 3. Partition-Specific TTL Isolation: Real-time Level-2 quote prices require aggressive expiry (5 seconds),
 *    whereas index constituent lists or corporate search mappings remain valid for hours or days.
 *    This configuration creates dedicated, size-capped cache partitions for each domain entity.
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@Configuration</b>: Marks this class as a Spring configuration bean providing bean definitions.</li>
 *   <li><b>@EnableCaching</b>: Enables Spring's annotation-driven cache abstraction ({@code @Cacheable}).</li>
 * </ul>
 */
@Configuration
@EnableCaching
public class CacheConfig {

    /**
     * Builds and registers the primary Spring {@link CacheManager} configured with custom Caffeine specifications.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@Bean</b>: Produces the application's central {@link CacheManager} bean autowired by Spring Cache.</li>
     * </ul>
     *
     * Partition Specifications:
     * <ul>
     *   <li><b>quotes</b>: Max 1,000 items | 5s TTL | Live quotes, LTP, circuit limits, and DOM order book.</li>
     *   <li><b>charts</b>: Max 500 items | 60s TTL | Intraday and multi-day candlestick time-series points.</li>
     *   <li><b>announcements</b>: Max 1,000 items | 300s (5m) TTL | Corporate regulatory filings & board notices.</li>
     *   <li><b>topten</b>: Max 10 items | 15s TTL | Market-wide top gainers, losers, and active volume stocks.</li>
     *   <li><b>yearwise</b>: Max 1,000 items | 24 hours TTL | Historical multi-period returns vs NIFTY 50.</li>
     *   <li><b>indices</b>: Max 1,000 items | 7 days TTL | Static index constituent listings.</li>
     *   <li><b>search</b>: Max 2,000 items | 1 hour TTL | Autocomplete fuzzy ticker search results.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * When 50 algorithmic strategies concurrently evaluate RELIANCE or TCS, only the first thread triggers
     * an upstream HTTP request to NSE; the remaining 49 threads receive instantaneous cached responses
     * directly from RAM with sub-microsecond latency.
     *
     * @return Fully configured {@link CacheManager} with custom Caffeine cache partitions.
     */
    @Bean
    public CacheManager cacheManager() {
        CaffeineCacheManager cacheManager = new CaffeineCacheManager();
        cacheManager.setCaffeine(Caffeine.newBuilder()
                .maximumSize(5000)
                .expireAfterWrite(60, TimeUnit.SECONDS)
                .recordStats());
        
        // Register distinct cache containers
        cacheManager.registerCustomCache("quotes",
                Caffeine.newBuilder().maximumSize(1000).expireAfterWrite(5, TimeUnit.SECONDS).build());
        cacheManager.registerCustomCache("charts",
                Caffeine.newBuilder().maximumSize(500).expireAfterWrite(60, TimeUnit.SECONDS).build());
        cacheManager.registerCustomCache("announcements",
                Caffeine.newBuilder().maximumSize(1000).expireAfterWrite(300, TimeUnit.SECONDS).build());
        cacheManager.registerCustomCache("topten",
                Caffeine.newBuilder().maximumSize(10).expireAfterWrite(15, TimeUnit.SECONDS).build());
        cacheManager.registerCustomCache("yearwise",
                Caffeine.newBuilder().maximumSize(1000).expireAfterWrite(24, TimeUnit.HOURS).build());
        cacheManager.registerCustomCache("indices",
                Caffeine.newBuilder().maximumSize(1000).expireAfterWrite(7, TimeUnit.DAYS).build());
        cacheManager.registerCustomCache("search",
                Caffeine.newBuilder().maximumSize(2000).expireAfterWrite(1, TimeUnit.HOURS).build());

        return cacheManager;
    }
}
