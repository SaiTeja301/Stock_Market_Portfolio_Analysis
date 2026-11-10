package com.stockmarket.gateway.client;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.stockmarket.gateway.dto.*;
import com.stockmarket.gateway.exception.NSEApiException;
import com.stockmarket.gateway.session.NSESessionManager;
import io.github.resilience4j.circuitbreaker.annotation.CircuitBreaker;
import io.github.resilience4j.ratelimiter.annotation.RateLimiter;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.HttpHeaders;
import org.springframework.stereotype.Component;
import org.springframework.web.reactive.function.client.WebClient;
import org.springframework.web.reactive.function.client.WebClientResponseException;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * ====================================================================================================
 *                 RESILIENT LOW-LEVEL NSE HTTP API CLIENT
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Core transport client responsible for executing low-level HTTP network calls against the National
 * Stock Exchange of India (NSE) public NextApi and GetQuoteApi infrastructure.
 *
 * Key Engineering Innovations:
 * 1. Akamai Bot-Defense Cookie Injection:
 *    NSE India employs strict Akamai anti-scraping perimeter security. Standard HTTP requests without
 *    valid cookies trigger immediate HTTP 403 Forbidden responses. This client works hand-in-hand with
 *    {@link NSESessionManager}, injecting active session cookies into the {@code Cookie} header of
 *    every outgoing request.
 * 2. Self-Healing Transparent Retry on Authentication Failure:
 *    If an upstream request encounters HTTP 401 or HTTP 403 (indicating cookie expiry), the client
 *    automatically catches the exception, commands {@link NSESessionManager} to force-harvest fresh
 *    cookies from the exchange homepage, updates the request headers, and performs a seamless single retry
 *    before throwing an error.
 * 3. Resilience4j Fault Tolerance & Circuit Breaking:
 *    Every public method is guarded by {@code @RateLimiter} (to avoid flooding upstream endpoints)
 *    and {@code @CircuitBreaker} (to prevent thread exhaustion when NSE experiences exchange downtime).
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@Component</b>: Declares this class as a Spring-managed singleton component in the IoC container.</li>
 *   <li><b>@Slf4j</b>: Generates an SLF4J Logger instance named {@code log} for execution tracing.</li>
 *   <li><b>@RequiredArgsConstructor</b>: Lombok annotation generating constructor injection for {@code nseWebClient},
 *       {@code sessionManager}, and {@code objectMapper}.</li>
 * </ul>
 */
@Slf4j
@Component
@RequiredArgsConstructor
public class NSEApiClient {

    private final WebClient nseWebClient;
    private final NSESessionManager sessionManager;
    private final ObjectMapper objectMapper;

    /**
     * Executes a low-level HTTP GET request with cookie injection, automated 401/403 recovery, and JSON parsing.
     *
     * Internal Mechanics:
     * 1. Retrieves latest Akamai session cookies from {@link NSESessionManager}.
     * 2. Issues blocking HTTP GET request via Spring's reactive {@link WebClient}.
     * 3. If upstream returns 401 Unauthorized or 403 Forbidden, force-refreshes session cookies and retries.
     * 4. Deserializes raw response JSON into Jackson's tree structure ({@link JsonNode}).
     *
     * Real-Time Trading Scenario:
     * Resilient high-availability connectivity. Market sessions can last hours; when exchange session tokens
     * expire mid-session, this method seamlessly recovers without failing downstream algo execution.
     *
     * @param uri The relative URI path on NSE's domain.
     * @return Parsed {@link JsonNode} root.
     * @throws NSEApiException If the network call fails or JSON cannot be parsed.
     */
    private JsonNode executeGet(String uri) {
        String cookies = sessionManager.getCookieHeader();
        try {
            String rawJson = nseWebClient.get()
                    .uri(uri)
                    .header(HttpHeaders.COOKIE, cookies)
                    .retrieve()
                    .bodyToMono(String.class)
                    .block();

            return objectMapper.readTree(rawJson);
        } catch (WebClientResponseException e) {
            if (e.getStatusCode().value() == 401 || e.getStatusCode().value() == 403) {
                log.warn("Upstream NSE returned HTTP {}. Force refreshing session cookies and retrying once...", e.getStatusCode().value());
                sessionManager.refreshSession(true);
                String refreshedCookies = sessionManager.getCookieHeader();
                String rawJson = nseWebClient.get()
                        .uri(uri)
                        .header(HttpHeaders.COOKIE, refreshedCookies)
                        .retrieve()
                        .bodyToMono(String.class)
                        .block();
                try {
                    return objectMapper.readTree(rawJson);
                } catch (Exception ex) {
                    throw new NSEApiException("Failed to parse JSON after retry: " + ex.getMessage(), ex, 502);
                }
            }
            throw new NSEApiException("NSE API returned HTTP " + e.getStatusCode().value() + ": " + e.getResponseBodyAsString(), e, e.getStatusCode().value());
        } catch (Exception e) {
            log.error("Error executing request against NSE endpoint [{}]: {}", uri, e.getMessage());
            throw new NSEApiException("NSE request failed: " + e.getMessage(), e, 502);
        }
    }

    /**
     * Fetches recent official corporate filings and board meeting notices from NSE GetQuoteApi.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@RateLimiter(name = "nseApiRateLimiter")</b>: Resilience4j rate limiter restricting
     *       concurrent request spikes to stay within safe NSE IP call budgets.</li>
     *   <li><b>@CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackCorporateAnnouncements")</b>:
     *       Trips open if upstream NSE error rate exceeds the 50% threshold, instantly executing
     *       {@link #fallbackCorporateAnnouncements} rather than hanging client threads.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * 48-Hour Binary Event Embargo pre-flight check. Systematic swing trading systems inspect this feed
     * to ensure no unexpected board actions or earnings releases threaten an open position.
     *
     * @param symbol NSE equity symbol.
     * @param noOfRecords Number of recent announcement records to fetch.
     * @return List of parsed {@link CorporateAnnouncementRecord} objects.
     */
    @RateLimiter(name = "nseApiRateLimiter")
    @CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackCorporateAnnouncements")
    public List<CorporateAnnouncementRecord> fetchCorporateAnnouncements(String symbol, int noOfRecords) {
        String encoded = URLEncoder.encode(symbol, StandardCharsets.UTF_8);
        String uri = String.format("/api/NextApi/apiClient/GetQuoteApi?functionName=getCorporateAnnouncement&symbol=%s&marketApiType=equities&noOfRecords=%d",
                encoded, noOfRecords);

        JsonNode root = executeGet(uri);
        List<CorporateAnnouncementRecord> list = new ArrayList<>();
        if (root.isArray()) {
            for (JsonNode node : root) {
                list.add(new CorporateAnnouncementRecord(
                        node.path("symbol").asText(symbol),
                        node.path("desc").asText(""),
                        node.path("an_dt").asText(""),
                        node.hasNonNull("attchmntFile") ? node.path("attchmntFile").asText() : null,
                        node.path("sm_name").asText(""),
                        node.path("sm_isin").asText(""),
                        node.path("attchmntText").asText(""),
                        node.path("hasXbrl").asBoolean(false)
                ));
            }
        }
        return list;
    }

    /**
     * Fetches real-time equity quote data, 52-week extremes, circuit limits, and 5-level DOM depth.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@RateLimiter(name = "nseApiRateLimiter")</b>: Prevents IP throttling on high-frequency quote polling.</li>
     *   <li><b>@CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackSymbolQuote")</b>:
     *       Fails gracefully to {@link #fallbackSymbolQuote} if NSE quote servers are unresponsive.</li>
     * </ul>
     *
     * Internal Mechanics:
     * Parses the composite JSON payload from NSE's {@code getSymbolData} API:
     * - {@code tradeInfo.lastPrice}: Real-time Last Traded Price (LTP).
     * - {@code metaData.averagePrice}: Volume Weighted Average Price (VWAP).
     * - {@code priceInfo.priceBand}: Upper and lower circuit limit bounds.
     * - {@code orderBook}: Top 5 buy bids and sell asks with quantities.
     *
     * Real-Time Trading Scenario:
     * High-frequency order placement pre-checks. Validates that the stock is not locked in upper or lower
     * circuit bands and checks orderbook bid-ask depth before firing orders.
     *
     * @param symbol NSE equity ticker symbol.
     * @return Fully populated {@link SymbolQuoteRecord}.
     */
    @RateLimiter(name = "nseApiRateLimiter")
    @CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackSymbolQuote")
    public SymbolQuoteRecord fetchSymbolData(String symbol) {
        String encoded = URLEncoder.encode(symbol, StandardCharsets.UTF_8);
        String uri = String.format("/api/NextApi/apiClient/GetQuoteApi?functionName=getSymbolData&marketType=N&series=EQ&symbol=%s", encoded);

        JsonNode root = executeGet(uri);
        JsonNode eqArray = root.path("equityResponse");
        if (!eqArray.isArray() || eqArray.isEmpty()) {
            throw new NSEApiException("Empty equityResponse from NSE for symbol: " + symbol, 404);
        }

        JsonNode eq = eqArray.get(0);
        JsonNode tradeInfo = eq.path("tradeInfo");
        JsonNode metaData = eq.path("metaData");
        JsonNode priceInfo = eq.path("priceInfo");

        double lowerCp = 0.0;
        double upperCp = 0.0;
        String priceBand = priceInfo.path("priceBand").asText("");
        if (priceBand.contains("-")) {
            String[] parts = priceBand.split("-");
            try {
                lowerCp = Double.parseDouble(parts[0].trim());
                upperCp = Double.parseDouble(parts[1].trim());
            } catch (NumberFormatException ignored) {}
        }

        OrderBookRecord orderBook = null;
        JsonNode obRaw = eq.path("orderBook");
        if (!obRaw.isMissingNode() && !obRaw.isNull() && obRaw.size() > 0) {
            List<OrderBookLevelRecord> levels = new ArrayList<>();
            for (int i = 1; i <= 5; i++) {
                double bp = obRaw.path("buyPrice" + i).asDouble(0.0);
                long bq = obRaw.path("buyQuantity" + i).asLong(0);
                double sp = obRaw.path("sellPrice" + i).asDouble(0.0);
                long sq = obRaw.path("sellQuantity" + i).asLong(0);
                levels.add(new OrderBookLevelRecord(i, bp, bq, sp, sq));
            }
            orderBook = new OrderBookRecord(
                    levels,
                    obRaw.path("totalBuyQuantity").asLong(0),
                    obRaw.path("totalSellQuantity").asLong(0)
            );
        }

        return new SymbolQuoteRecord(
                symbol,
                metaData.path("series").asText("EQ"),
                tradeInfo.path("lastPrice").asDouble(0.0),
                metaData.path("change").asDouble(0.0),
                metaData.path("pChange").asDouble(0.0),
                metaData.path("previousClose").asDouble(0.0),
                metaData.path("open").asDouble(0.0),
                metaData.path("closePrice").asDouble(0.0),
                metaData.path("averagePrice").asDouble(0.0),
                lowerCp,
                upperCp,
                priceInfo.path("yearHigh").asDouble(0.0),
                priceInfo.path("yearLow").asDouble(0.0),
                orderBook
        );
    }

    /**
     * Fetches historical and intraday time-series chart data points from NSE GetQuoteApi.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@RateLimiter(name = "nseApiRateLimiter")</b>: Throttle protection on heavy chart endpoints.</li>
     *   <li><b>@CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackSymbolChart")</b>:
     *       Resilience4j fallback to empty chart structure if NSE chart engine fails.</li>
     * </ul>
     *
     * Internal Mechanics:
     * Automatically appends the required NSE identifier suffix ("EQN") if omitted. Parses the
     * multi-dimensional array {@code grapthData} extracting timestamps, closing values, and volumes.
     *
     * Real-Time Trading Scenario:
     * Quantitative regression modeling. Feeds time-series vectors into the 100-period OLS Regression Channel.
     *
     * @param symbol NSE equity symbol.
     * @param days Timeframe window ("1D", "1W", "1M", "1Y").
     * @return {@link SymbolChartRecord} containing timestamped chart data points.
     */
    @RateLimiter(name = "nseApiRateLimiter")
    @CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackSymbolChart")
    public SymbolChartRecord fetchSymbolChartData(String symbol, String days) {
        String identifier = symbol.endsWith("EQN") ? symbol : symbol + "EQN";
        String encoded = URLEncoder.encode(identifier, StandardCharsets.UTF_8);
        String uri = String.format("/api/NextApi/apiClient/GetQuoteApi?functionName=getSymbolChartData&symbol=%s&days=%s",
                encoded, days != null ? days : "1D");

        JsonNode root = executeGet(uri);
        List<ChartPointRecord> points = new ArrayList<>();
        JsonNode graphData = root.path("grapthData");
        if (graphData.isArray()) {
            for (JsonNode pt : graphData) {
                if (pt.isArray() && pt.size() >= 5) {
                    points.add(new ChartPointRecord(
                            pt.get(0).asLong(),
                            pt.get(1).asDouble(),
                            pt.get(2).asText(""),
                            pt.get(3).asDouble(),
                            pt.get(4).asDouble()
                    ));
                }
            }
        }

        return new SymbolChartRecord(
                root.path("identifier").asText(identifier),
                root.path("name").asText(symbol),
                root.path("closePrice").asDouble(0.0),
                points
        );
    }

    /**
     * Fetches multi-period historical returns (1D, 1W, 1M, 1Y, 3Y, 5Y) and benchmark index comparison.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@RateLimiter(name = "nseApiRateLimiter")</b>: Guard against rapid consecutive requests.</li>
     *   <li><b>@CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackYearwise")</b>:
     *       Fallback handler returning zeroed baseline record during exchange service degradations.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Multi-period relative momentum scoring. Identifies institutional stocks with sustained alpha vs NIFTY 50.
     *
     * @param symbol NSE equity ticker symbol.
     * @return Populated {@link YearwisePerformanceRecord}.
     */
    @RateLimiter(name = "nseApiRateLimiter")
    @CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackYearwise")
    public YearwisePerformanceRecord fetchYearwiseData(String symbol) {
        String identifier = symbol.endsWith("EQN") ? symbol : symbol + "EQN";
        String encoded = URLEncoder.encode(identifier, StandardCharsets.UTF_8);
        String uri = String.format("/api/NextApi/apiClient/GetQuoteApi?functionName=getYearwiseData&symbol=%s", encoded);

        JsonNode root = executeGet(uri);
        JsonNode item = root.isArray() && !root.isEmpty() ? root.get(0) : root;

        return new YearwisePerformanceRecord(
                item.path("yesterday_chng_per").asDouble(0.0),
                item.path("one_week_chng_per").asDouble(0.0),
                item.path("one_month_chng_per").asDouble(0.0),
                item.path("three_month_chng_per").asDouble(0.0),
                item.path("six_month_chng_per").asDouble(0.0),
                item.path("one_year_chng_per").asDouble(0.0),
                item.path("two_year_chng_per").asDouble(0.0),
                item.path("three_year_chng_per").asDouble(0.0),
                item.path("five_year_chng_per").asDouble(0.0),
                item.path("index_name").asText("NIFTY 50"),
                item.path("index_one_year_chng_per").asDouble(0.0)
        );
    }

    /**
     * Resolves the list of sectoral and benchmark indices containing the specified equity symbol.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@RateLimiter(name = "nseApiRateLimiter")</b>: Rate limiter protecting NSE index endpoints.</li>
     *   <li><b>@CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackIndexList")</b>:
     *       Circuit breaker returning empty list on failure.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Sectoral contagion and index rebalancing surveillance.
     *
     * @param symbol NSE equity symbol.
     * @return List of index names.
     */
    @RateLimiter(name = "nseApiRateLimiter")
    @CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackIndexList")
    public List<String> fetchIndexList(String symbol) {
        String encoded = URLEncoder.encode(symbol, StandardCharsets.UTF_8);
        String uri = String.format("/api/NextApi/apiClient/GetQuoteApi?functionName=getIndexList&symbol=%s", encoded);

        JsonNode root = executeGet(uri);
        List<String> list = new ArrayList<>();
        if (root.isArray()) {
            for (JsonNode item : root) {
                list.add(item.asText());
            }
        }
        return list;
    }

    /**
     * Fetches market-wide top 10 gainers, top 10 losers, most active equities, and volume spurts.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@RateLimiter(name = "nseApiRateLimiter")</b>: Rate limits market-wide data queries.</li>
     *   <li><b>@CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackTopTen")</b>:
     *       Returns empty fallback structure if market breadth services are unavailable.</li>
     *   <li><b>@SuppressWarnings("unchecked")</b>: Suppresses unchecked cast warnings during Jackson
     *       raw Map conversions for variable-length active volume and value lists.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Market breadth and intraday momentum screening. Systematic strategies scan top gainers and volume
     * spurts at market open to trigger momentum breakouts.
     *
     * @return {@link TopTenStocksRecord} containing gainers, losers, and volume leaders.
     */
    @RateLimiter(name = "nseApiRateLimiter")
    @CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackTopTen")
    @SuppressWarnings("unchecked")
    public TopTenStocksRecord fetchTopTenStocks() {
        String uri = "/api/NextApi/apiClient/GetQuoteApi?functionName=getTopTenStock";
        JsonNode root = executeGet(uri);

        List<TopGainerLoserRecord> gainers = parseGainersLosers(root.path("topGainers"));
        List<TopGainerLoserRecord> losers = parseGainersLosers(root.path("topLoosers"));

        List<Map<String, Object>> activeVal = objectMapper.convertValue(root.path("mostActiveValue"), List.class);
        List<Map<String, Object>> activeVol = objectMapper.convertValue(root.path("mostActiveVolume"), List.class);
        List<Map<String, Object>> spurts = objectMapper.convertValue(root.path("volumeSpurtsValue"), List.class);

        return new TopTenStocksRecord(
                root.path("timestamp").asText(""),
                gainers != null ? gainers : Collections.emptyList(),
                losers != null ? losers : Collections.emptyList(),
                activeVal != null ? activeVal : Collections.emptyList(),
                activeVol != null ? activeVol : Collections.emptyList(),
                spurts != null ? spurts : Collections.emptyList()
        );
    }

    /**
     * Helper parser transforming JSON arrays into typed {@link TopGainerLoserRecord} lists.
     *
     * @param array Jackson JSON array node.
     * @return List of {@link TopGainerLoserRecord} domain objects.
     */
    private List<TopGainerLoserRecord> parseGainersLosers(JsonNode array) {
        if (!array.isArray()) return Collections.emptyList();
        List<TopGainerLoserRecord> list = new ArrayList<>();
        for (JsonNode g : array) {
            list.add(new TopGainerLoserRecord(
                    g.path("symbol").asText(""),
                    g.path("series").asText("EQ"),
                    g.path("lastPrice").asDouble(0.0),
                    g.path("change").asDouble(0.0),
                    g.path("pChange").asDouble(0.0),
                    g.path("totalTradedVolume").asLong(0),
                    g.path("totalTradedValue").asDouble(0.0)
            ));
        }
        return list;
    }

    /**
     * Executes fuzzy search against NSE equity database for ticker and company resolution.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@RateLimiter(name = "nseApiRateLimiter")</b>: Throttles search-as-you-type input spikes.</li>
     *   <li><b>@CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackSearch")</b>:
     *       Circuit breaker returning empty search results on exchange downtime.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * User query autocomplete and ISIN / Symbol resolution.
     *
     * @param query Search string.
     * @return List of matching {@link SearchResultRecord} objects.
     */
    @RateLimiter(name = "nseApiRateLimiter")
    @CircuitBreaker(name = "nseApiCircuitBreaker", fallbackMethod = "fallbackSearch")
    public List<SearchResultRecord> fetchGlobalSearchEquity(String query) {
        String encoded = URLEncoder.encode(query, StandardCharsets.UTF_8);
        String uri = String.format("/api/NextApi/globalSearch/equity?symbol=%s", encoded);

        JsonNode root = executeGet(uri);
        List<SearchResultRecord> list = new ArrayList<>();
        JsonNode data = root.path("data");
        if (data.isArray()) {
            for (JsonNode item : data) {
                list.add(new SearchResultRecord(
                        item.path("symbol").asText(""),
                        item.path("companyName").asText(""),
                        item.path("series").asText("EQ"),
                        item.path("segment").asText(""),
                        item.path("lastPrice").asDouble(0.0),
                        item.path("url").asText("")
                ));
            }
        }
        return list;
    }

    // ================================================================================================
    //                               RESILIENCE4J CIRCUIT BREAKER FALLBACKS
    // ================================================================================================

    /**
     * Graceful fallback for corporate announcements when the circuit breaker trips or times out.
     *
     * @param symbol Equity ticker symbol.
     * @param noOfRecords Requested record count.
     * @param t Root cause exception.
     * @return Empty list preventing null pointer exceptions in callers.
     */
    public List<CorporateAnnouncementRecord> fallbackCorporateAnnouncements(String symbol, int noOfRecords, Throwable t) {
        log.warn("Fallback triggered for announcements symbol [{}]: {}", symbol, t.getMessage());
        return Collections.emptyList();
    }

    /**
     * Graceful fallback for symbol quote data when the circuit breaker is OPEN.
     *
     * @param symbol Target ticker symbol.
     * @param t Root cause exception.
     * @return Zeroed {@link SymbolQuoteRecord} preventing cascading service failure.
     */
    public SymbolQuoteRecord fallbackSymbolQuote(String symbol, Throwable t) {
        log.warn("Fallback triggered for symbol quote [{}]: {}", symbol, t.getMessage());
        return new SymbolQuoteRecord(symbol, "EQ", 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, null);
    }

    /**
     * Graceful fallback for historical chart data when the circuit breaker is OPEN.
     *
     * @param symbol Target ticker symbol.
     * @param days Requested timeframe.
     * @param t Root cause exception.
     * @return Empty {@link SymbolChartRecord}.
     */
    public SymbolChartRecord fallbackSymbolChart(String symbol, String days, Throwable t) {
        log.warn("Fallback triggered for chart [{}]: {}", symbol, t.getMessage());
        return new SymbolChartRecord(symbol, symbol, 0.0, Collections.emptyList());
    }

    /**
     * Graceful fallback for yearwise performance metrics when the circuit breaker is OPEN.
     *
     * @param symbol Target ticker symbol.
     * @param t Root cause exception.
     * @return Default zero-initialized {@link YearwisePerformanceRecord}.
     */
    public YearwisePerformanceRecord fallbackYearwise(String symbol, Throwable t) {
        log.warn("Fallback triggered for yearwise [{}]: {}", symbol, t.getMessage());
        return new YearwisePerformanceRecord(0,0,0,0,0,0,0,0,0,"NIFTY 50",0);
    }

    /**
     * Graceful fallback for index constituent list when the circuit breaker is OPEN.
     *
     * @param symbol Target ticker symbol.
     * @param t Root cause exception.
     * @return Empty list.
     */
    public List<String> fallbackIndexList(String symbol, Throwable t) {
        log.warn("Fallback triggered for index list [{}]: {}", symbol, t.getMessage());
        return Collections.emptyList();
    }

    /**
     * Graceful fallback for top ten stocks when the circuit breaker is OPEN.
     *
     * @param t Root cause exception.
     * @return Empty {@link TopTenStocksRecord}.
     */
    public TopTenStocksRecord fallbackTopTen(Throwable t) {
        log.warn("Fallback triggered for top ten: {}", t.getMessage());
        return new TopTenStocksRecord("", Collections.emptyList(), Collections.emptyList(), Collections.emptyList(), Collections.emptyList(), Collections.emptyList());
    }

    /**
     * Graceful fallback for equity search queries when the circuit breaker is OPEN.
     *
     * @param query User search string.
     * @param t Root cause exception.
     * @return Empty list.
     */
    public List<SearchResultRecord> fallbackSearch(String query, Throwable t) {
        log.warn("Fallback triggered for search [{}]: {}", query, t.getMessage());
        return Collections.emptyList();
    }
}
