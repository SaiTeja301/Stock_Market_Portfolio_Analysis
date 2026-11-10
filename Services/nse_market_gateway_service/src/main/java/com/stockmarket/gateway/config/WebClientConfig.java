package com.stockmarket.gateway.config;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.http.HttpHeaders;
import org.springframework.http.client.reactive.ReactorClientHttpConnector;
import org.springframework.web.reactive.function.client.ExchangeStrategies;
import org.springframework.web.reactive.function.client.WebClient;
import reactor.netty.http.client.HttpClient;

import java.time.Duration;

/**
 * ====================================================================================================
 *                 REACTIVE WEBCLIENT TRANSPORT CONFIGURATION
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Configures the non-blocking Spring Reactive {@link WebClient} powering all outbound network
 * communication to the National Stock Exchange of India (NSE).
 *
 * Key Transport Configurations:
 * 1. Reactor Netty HTTP Connector:
 *    Underpinned by high-throughput Reactor Netty with configurable socket read and response timeouts
 *    to prevent thread hanging during exchange network hiccups.
 * 2. 16MB In-Memory Codec Buffer:
 *    Spring WebClient defaults to a 256KB buffer limit. NSE market-wide payloads (e.g. historical
 *    intraday charts or top ten market scans) can reach 2MB to 8MB in raw JSON. The codec configuration
 *    expands the max in-memory buffer to 16MB ({@code 16 * 1024 * 1024}), eliminating
 *    {@code DataBufferLimitException: Exceeded limit on max bytes to buffer}.
 * 3. Browser Emulation Headers:
 *    NSE NextApi endpoints mandate exact browser headers. The client automatically applies:
 *    - {@code User-Agent}: Chrome 128 on Windows 64-bit.
 *    - {@code Referer}: "https://www.nseindia.com/".
 *    - {@code X-Requested-With}: "XMLHttpRequest" (identifies requests as valid AJAX calls).
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@Configuration</b>: Declares this class as a Spring configuration bean.</li>
 * </ul>
 */
@Configuration
public class WebClientConfig {

    @Value("${nse.api.base-url:https://www.nseindia.com}")
    private String baseUrl;

    @Value("${nse.api.connect-timeout-ms:6000}")
    private int connectTimeoutMs;

    @Value("${nse.api.read-timeout-ms:6000}")
    private int readTimeoutMs;

    /**
     * Creates and exposes the configured {@link WebClient} bean used by {@link com.stockmarket.gateway.client.NSEApiClient}.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@Bean</b>: Registers the singleton WebClient instance in the Spring IoC container.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Ensures low-latency, non-blocking outbound HTTP requests with full browser emulation,
     * allowing high-frequency algorithmic queries to execute without triggering WAF blocks.
     *
     * @return Fully configured {@link WebClient} instance.
     */
    @Bean
    public WebClient nseWebClient() {
        HttpClient httpClient = HttpClient.create()
                .responseTimeout(Duration.ofMillis(readTimeoutMs))
                .followRedirect(true);

        ExchangeStrategies strategies = ExchangeStrategies.builder()
                .codecs(configurer -> configurer.defaultCodecs().maxInMemorySize(16 * 1024 * 1024))
                .build();

        return WebClient.builder()
                .baseUrl(baseUrl)
                .clientConnector(new ReactorClientHttpConnector(httpClient))
                .exchangeStrategies(strategies)
                .defaultHeader(HttpHeaders.USER_AGENT, "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")
                .defaultHeader(HttpHeaders.REFERER, "https://www.nseindia.com/")
                .defaultHeader(HttpHeaders.ACCEPT, "*/*")
                .defaultHeader(HttpHeaders.ACCEPT_LANGUAGE, "en-US,en;q=0.9")
                .defaultHeader("X-Requested-With", "XMLHttpRequest")
                .build();
    }
}
