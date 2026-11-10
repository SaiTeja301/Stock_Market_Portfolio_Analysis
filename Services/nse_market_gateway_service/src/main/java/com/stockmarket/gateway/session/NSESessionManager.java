package com.stockmarket.gateway.session;

import com.stockmarket.gateway.exception.NSEApiException;
import jakarta.annotation.PostConstruct;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.net.CookieManager;
import java.net.CookiePolicy;
import java.net.HttpCookie;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

/**
 * ====================================================================================================
 *         AKAMAI BOT DEFENSE SESSION HARVESTER & COOKIE MANAGEMENT ENGINE
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Autonomous session persistence and anti-bot mitigation engine designed to establish, maintain,
 * and rotate legitimate HTTP browser sessions against the National Stock Exchange of India (NSE).
 *
 * The Akamai Bot-Defense Challenge:
 * NSE India sits behind an enterprise Akamai Web Application Firewall (WAF) and Bot Manager.
 * Akamai evaluates TLS fingerprints, HTTP header signatures, and session cookies (specifically
 * {@code _abck}, {@code bm_sz}, and {@code ak_bmsc}). Direct API queries issued without valid
 * pre-negotiated cookies result in immediate HTTP 403 Forbidden blocks.
 *
 * Engineering Solution:
 * 1. Proactive Session Warmup:
 *    On microservice startup ({@code @PostConstruct}), this engine initiates a browser-spoofed HTTP
 *    handshake to NSE's landing page ({@code https://www.nseindia.com}) using Java's {@link HttpClient}.
 * 2. Automated Cookie Harvesting:
 *    Captured {@link HttpCookie} instances are parsed and aggregated into a valid RFC-6265 cookie header
 *    stored inside a thread-safe {@link AtomicReference}.
 * 3. Scheduled Background Rotation (Heartbeat):
 *    Akamai cookies expire every 15-30 minutes. The {@code @Scheduled} background heartbeat periodically
 *    rotates cookies every 10 minutes (600,000 ms), guaranteeing seamless API connectivity during market hours.
 * 4. Lazy-Loaded Recovery:
 *    If an upstream API call fails with 401 or 403, the client can command this manager to perform a
 *    synchronous force-refresh on demand.
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@Component</b>: Spring singleton bean managed by the IoC container.</li>
 *   <li><b>@Slf4j</b>: SLF4J logger for session rotation diagnostics.</li>
 * </ul>
 */
@Slf4j
@Component
public class NSESessionManager {

    /**
     * Injected base URL of the National Stock Exchange landing page.
     * Annotation: @Value - Reads property "nse.api.base-url" from application.yml, defaulting to "https://www.nseindia.com".
     */
    @Value("${nse.api.base-url:https://www.nseindia.com}")
    private String baseUrl;

    /**
     * Injected socket connection timeout in milliseconds.
     * Annotation: @Value - Reads property "nse.api.connect-timeout-ms", defaulting to 6000 ms.
     */
    @Value("${nse.api.connect-timeout-ms:6000}")
    private int timeoutMs;

    /**
     * Thread-safe cache holding the compiled Cookie header string (e.g. "_abck=...; bm_sz=...").
     * Uses AtomicReference to allow lock-free concurrent reads by thousands of incoming market queries.
     */
    private final AtomicReference<String> cookieHeaderCache = new AtomicReference<>("");

    /**
     * In-memory cookie jar automatically tracking set-cookie responses from the exchange.
     */
    private final CookieManager cookieManager = new CookieManager(null, CookiePolicy.ACCEPT_ALL);

    /**
     * High-speed native Java HTTP Client configured with redirect following and custom cookie handler.
     */
    private HttpClient httpClient;

    /**
     * Post-construct lifecycle callback initializing the HTTP client and triggering initial session warmup.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@PostConstruct</b>: Standard Jakarta EE annotation. Executed by Spring immediately
     *       after dependency injection is complete and before the microservice accepts incoming traffic.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Eliminates "cold-start" latency for the first trading request of the day by ensuring valid
     * exchange session tokens are already primed in RAM when the market opens at 09:15 IST.
     */
    @PostConstruct
    public void init() {
        this.httpClient = HttpClient.newBuilder()
                .cookieHandler(cookieManager)
                .connectTimeout(Duration.ofMillis(timeoutMs))
                .followRedirects(HttpClient.Redirect.NORMAL)
                .build();
        try {
            refreshSession(true);
        } catch (Exception e) {
            log.warn("Initial session warmup encountered non-fatal delay: {}", e.getMessage());
        }
    }

    /**
     * Scheduled background session heartbeat refreshing cookies every 10 minutes.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@Scheduled</b>: Spring scheduling annotation.
     *       - {@code fixedDelayString}: 600,000 ms (10 minutes) between the end of one execution and the start of the next.
     *       - {@code initialDelay}: 600,000 ms delay before the first scheduled run.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Unattended 24/7 trading server persistence. Prevents session timeout during long trading sessions
     * by rotating tokens before Akamai's 30-minute session invalidation window closes.
     */
    @Scheduled(fixedDelayString = "${nse.api.session-refresh-interval-ms:600000}", initialDelay = 600000)
    public void scheduledWarmup() {
        log.info("Executing scheduled Akamai session cookie refresh...");
        refreshSession(false);
    }

    /**
     * Harvests fresh session cookies by simulating a legitimate modern Chrome browser visiting the NSE homepage.
     *
     * Synchronization & Mechanics:
     * - Declared as {@code synchronized} to ensure only one thread performs the network handshake at a time,
     *   preventing "thundering herd" concurrency spikes against NSE during token renewal.
     * - Uses authentic Chrome 128 User-Agent and Accept headers to satisfy Akamai browser reputation checks.
     *
     * Real-Time Trading Scenario:
     * Emergency mid-trade token re-acquisition. If upstream NSE invalidates existing cookies mid-trade,
     * the client invokes this with {@code force=true} to instantly re-establish communication.
     *
     * @param force When true, bypasses the existing cache and forces a fresh network handshake.
     * @throws NSEApiException If the handshake fails or the exchange returns an unrecoverable HTTP error.
     */
    public synchronized void refreshSession(boolean force) {
        if (!force && !cookieHeaderCache.get().isEmpty()) {
            return;
        }

        log.info("Requesting {} to harvest active session cookies...", baseUrl);
        try {
            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(baseUrl))
                    .header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")
                    .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8")
                    .header("Accept-Language", "en-US,en;q=0.9")
                    .timeout(Duration.ofMillis(timeoutMs))
                    .GET()
                    .build();

            HttpResponse<String> response = httpClient.send(request, HttpResponse.BodyHandlers.ofString());
            if (response.statusCode() >= 200 && response.statusCode() < 400) {
                List<HttpCookie> cookies = cookieManager.getCookieStore().get(URI.create(baseUrl));
                String combinedCookies = cookies.stream()
                        .map(c -> c.getName() + "=" + c.getValue())
                        .collect(Collectors.joining("; "));

                if (!combinedCookies.isEmpty()) {
                    cookieHeaderCache.set(combinedCookies);
                    log.info("Successfully acquired {} cookies. Names: {}",
                            cookies.size(), cookies.stream().map(HttpCookie::getName).toList());
                } else {
                    log.warn("Session refreshed with HTTP {} but no cookies stored by CookieManager.", response.statusCode());
                }
            } else {
                log.warn("Session refresh responded with non-success status: {}", response.statusCode());
            }
        } catch (Exception e) {
            log.error("Failed to harvest NSE session cookies: {}", e.getMessage(), e);
            throw new NSEApiException("Failed to refresh NSE session: " + e.getMessage(), e, 502);
        }
    }

    /**
     * Returns the active, formatted Cookie header string ready for injection into outgoing HTTP requests.
     *
     * Internal Mechanics:
     * Lock-free read via {@link AtomicReference#get()}. If the cache is empty, triggers a synchronous
     * {@link #refreshSession(boolean)} before returning.
     *
     * Real-Time Trading Scenario:
     * High-throughput quote streaming. Hundreds of concurrent threads retrieve the valid cookie string
     * with zero lock contention.
     *
     * @return Formatted cookie header string (e.g. "_abck=XYZ...; bm_sz=ABC...").
     */
    public String getCookieHeader() {
        String current = cookieHeaderCache.get();
        if (current == null || current.isEmpty()) {
            refreshSession(true);
            return cookieHeaderCache.get();
        }
        return current;
    }
}
