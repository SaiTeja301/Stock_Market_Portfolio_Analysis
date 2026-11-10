package com.stockmarket.gateway.controller;

import com.stockmarket.gateway.dto.*;
import com.stockmarket.gateway.service.NSEMarketGatewayService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

/**
 * ====================================================================================================
 *                 NSE MARKET GATEWAY REST CONTROLLER: PUBLIC MARKET FACADE
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Exposes high-throughput, low-latency REST endpoints delivering real-time equity market data
 * fetched directly from the National Stock Exchange of India (NSE).
 *
 * It serves as the primary ingress facade for:
 * 1. Live Level-2 Market Depth (5-level DOM order book, LTP, VWAP, circuit bands).
 * 2. Multi-timeframe historical candlestick time-series chart data.
 * 3. Official corporate regulatory filings and board meeting intimations.
 * 4. Market-wide top gainers, losers, and volume leaders.
 * 5. Relative alpha performance tracking against benchmark indices (NIFTY 50).
 * 6. Global symbol fuzzy search and ticker resolution.
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@RestController</b>: Combines Spring's {@code @Controller} and {@code @ResponseBody}.
 *       Instructs Spring MVC to route HTTP requests directly to handler methods and automatically
 *       serialize domain return records into JSON HTTP response bodies using Jackson ObjectMapper.</li>
 *   <li><b>@RequestMapping("/api/v1/market")</b>: Establishes the authoritative base URI path prefix
 *       for all market data endpoints, enforcing standard API versioning (/api/v1/market) for
 *       downstream trading microservices and algorithmic execution engines.</li>
 *   <li><b>@RequiredArgsConstructor</b>: Lombok annotation generating a constructor for all {@code final}
 *       fields ({@code gatewayService}). Enforces immutable, compile-time verified constructor injection
 *       as recommended by the Spring engineering team.</li>
 *   <li><b>@Slf4j</b>: Generates an SLF4J Logger instance named {@code log} for structured, asynchronous
 *       request tracing and audit logging without JVM overhead.</li>
 * </ul>
 */
@Slf4j
@RestController
@RequestMapping("/api/v1/market")
@RequiredArgsConstructor
public class NSEMarketGatewayController {

    private final NSEMarketGatewayService gatewayService;

    /**
     * Fetches real-time sanitized quote data for an NSE equity, including Level-2 DOM order book depth.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/quote/{symbol}")</b>: Maps HTTP GET requests to "/api/v1/market/quote/{symbol}".
     *       GET is idempotent and safe for high-frequency polling by market ticker widgets.</li>
     *   <li><b>@PathVariable String symbol</b>: Extracts the equity ticker symbol (e.g. "RELIANCE", "TCS", "INFY")
     *       from the URI path segment directly into the method argument.</li>
     * </ul>
     *
     * Internal Mechanics:
     * Dispatches to {@link NSEMarketGatewayService#getSymbolQuote(String)}, which checks the Caffeine
     * in-memory L1 cache (10s TTL). On a cache miss, it calls {@code NSEApiClient} to query NSE's quote API,
     * injects Akamai session cookies, applies Resilience4j rate limiting, and deserializes the JSON response.
     *
     * Real-Time Trading Scenario:
     * Pre-trade execution verification. Algorithmic trade execution engines call this endpoint to inspect
     * live bid-ask spread, verify upper and lower circuit limits (ensuring the stock is not locked), and
     * confirm that current price is trading near VWAP before placing institutional block orders.
     *
     * @param symbol NSE equity ticker symbol (e.g. "RELIANCE", "TATAMOTORS", "HDFCBANK").
     * @return HTTP 200 OK containing {@link SymbolQuoteRecord} with LTP, VWAP, circuit limits, and order book.
     */
    @GetMapping("/quote/{symbol}")
    public ResponseEntity<SymbolQuoteRecord> getQuote(@PathVariable String symbol) {
        log.info("REST request to fetch live quote for symbol: [{}]", symbol);
        return ResponseEntity.ok(gatewayService.getSymbolQuote(symbol));
    }

    /**
     * Fetches multi-timeframe historical price chart data points for candlestick and regression channel rendering.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/charts/{symbol}")</b>: Maps HTTP GET requests to "/api/v1/market/charts/{symbol}".</li>
     *   <li><b>@PathVariable String symbol</b>: The target equity ticker symbol.</li>
     *   <li><b>@RequestParam(defaultValue = "1D") String days</b>: Optional query parameter specifying the
     *       chart timeframe resolution (accepted values: "1D", "1W", "1M", "1Y"). Defaults to intraday ("1D").</li>
     * </ul>
     *
     * Internal Mechanics:
     * Queries {@link NSEMarketGatewayService#getSymbolChart(String, String)}. Cached under key
     * {@code "{SYMBOL}:{DAYS}"} with a 15-second TTL. The data points contain epoch timestamps and
     * closing prices used to construct technical indicators.
     *
     * Real-Time Trading Scenario:
     * Technical charting and algorithmic linear regression channels. Quantitative strategies pull this time-series
     * to compute 100-period Ordinary Least Squares (OLS) regression channels and identify mean-reversion buy zones.
     *
     * @param symbol NSE equity ticker symbol (e.g. "TCS", "INFY").
     * @param days Timeframe window ("1D", "1W", "1M", "1Y"). Default is "1D".
     * @return HTTP 200 OK containing {@link SymbolChartRecord} with timestamped price points.
     */
    @GetMapping("/charts/{symbol}")
    public ResponseEntity<SymbolChartRecord> getChart(
            @PathVariable String symbol,
            @RequestParam(defaultValue = "1D") String days) {
        log.info("REST request to fetch chart data for symbol: [{}] with timeframe: [{}]", symbol, days);
        return ResponseEntity.ok(gatewayService.getSymbolChart(symbol, days));
    }

    /**
     * Retrieves recent official corporate announcements, regulatory filings, and board meeting notices.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/announcements/{symbol}")</b>: Maps HTTP GET requests to "/api/v1/market/announcements/{symbol}".</li>
     *   <li><b>@PathVariable String symbol</b>: The target equity ticker symbol.</li>
     *   <li><b>@RequestParam(defaultValue = "5") int limit</b>: Maximum number of recent corporate filings
     *       to return. Defaults to 5 records.</li>
     * </ul>
     *
     * Internal Mechanics:
     * Cached with a 60-second TTL in Caffeine. Queries NSE's corporate announcement feed, parsing publication
     * timestamps, description text, ISIN, and PDF attachment links.
     *
     * Real-Time Trading Scenario:
     * 48-Hour Binary Event Risk Embargo Check. Systematic swing trading systems inspect this endpoint
     * prior to entering positions to ensure no quarterly earnings release, dividend decision, or regulatory
     * inquiry is scheduled within the next 48 hours, preventing binary volatility shocks.
     *
     * @param symbol NSE equity ticker symbol.
     * @param limit Maximum number of announcements to return (default: 5).
     * @return HTTP 200 OK containing a list of {@link CorporateAnnouncementRecord} objects.
     */
    @GetMapping("/announcements/{symbol}")
    public ResponseEntity<List<CorporateAnnouncementRecord>> getAnnouncements(
            @PathVariable String symbol,
            @RequestParam(defaultValue = "5") int limit) {
        log.info("REST request to fetch corporate announcements for symbol: [{}] (limit: {})", symbol, limit);
        return ResponseEntity.ok(gatewayService.getCorporateAnnouncements(symbol, limit));
    }

    /**
     * Retrieves the market-wide top 10 gainers, top 10 losers, and most active equities by volume/turnover.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@GetMapping("/top-ten")</b>: Maps HTTP GET requests to "/api/v1/market/top-ten".</li>
     * </ul>
     *
     * Internal Mechanics:
     * Cached globally with a 30-second TTL to shield upstream NSE servers from redundant market-wide scans.
     * Aggregates percentage changes, trading volumes, and turnover values across the NIFTY universe.
     *
     * Real-Time Trading Scenario:
     * Morning momentum screening and intraday sector rotation surveillance. Day traders and automated scanners
     * poll this endpoint at market open (09:15 IST) to identify high-relative-volume leaders for momentum breakouts.
     *
     * @return HTTP 200 OK containing {@link TopTenStocksRecord} with gainers, losers, and active stocks.
     */
    @GetMapping("/top-ten")
    public ResponseEntity<TopTenStocksRecord> getTopTenStocks() {
        log.info("REST request to fetch global Top 10 gainers, losers, and most active equities");
        return ResponseEntity.ok(gatewayService.getTopTenStocks());
    }

    /**
     * Retrieves multi-period performance metrics (1D, 1W, 1M, 1Y, 3Y, 5Y) and benchmark comparison against NIFTY 50.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/yearwise/{symbol}")</b>: Maps HTTP GET requests to "/api/v1/market/yearwise/{symbol}".</li>
     *   <li><b>@PathVariable String symbol</b>: The target equity ticker symbol.</li>
     * </ul>
     *
     * Internal Mechanics:
     * Cached under key {@code "{SYMBOL}"} with a 60-second TTL. Fetches historical period closing prices,
     * calculates compounded returns, and computes the relative alpha generated above the NIFTY 50 index.
     *
     * Real-Time Trading Scenario:
     * Quantitative relative strength screening. Long-term trend-following models evaluate whether an equity
     * is outperforming the broader index across multiple rolling lookback windows before authorizing capital allocation.
     *
     * @param symbol NSE equity ticker symbol (e.g. "RELIANCE", "BAJFINANCE").
     * @return HTTP 200 OK containing {@link YearwisePerformanceRecord} with multi-period returns and index comparison.
     */
    @GetMapping("/yearwise/{symbol}")
    public ResponseEntity<YearwisePerformanceRecord> getYearwise(@PathVariable String symbol) {
        log.info("REST request to fetch yearwise multi-period performance for symbol: [{}]", symbol);
        return ResponseEntity.ok(gatewayService.getYearwisePerformance(symbol));
    }

    /**
     * Resolves the list of sectoral and thematic benchmark indices that include the specified equity.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/indices/{symbol}")</b>: Maps HTTP GET requests to "/api/v1/market/indices/{symbol}".</li>
     *   <li><b>@PathVariable String symbol</b>: Target equity ticker symbol.</li>
     * </ul>
     *
     * Internal Mechanics:
     * Cached with a 5-minute TTL. Maps ticker membership across NSE index definitions
     * (e.g. "NIFTY 50", "NIFTY IT", "NIFTY BANK", "NIFTY 500").
     *
     * Real-Time Trading Scenario:
     * Sector contagion and ETF basket arbitrage analysis. When a sector index experiences a sharp macro selloff,
     * risk systems query this endpoint to identify all portfolio stocks exposed to that index.
     *
     * @param symbol NSE equity ticker symbol.
     * @return HTTP 200 OK containing a list of index names.
     */
    @GetMapping("/indices/{symbol}")
    public ResponseEntity<List<String>> getIndices(@PathVariable String symbol) {
        log.info("REST request to fetch index membership for symbol: [{}]", symbol);
        return ResponseEntity.ok(gatewayService.getIndexList(symbol));
    }

    /**
     * Executes fuzzy search across NSE equities by ticker symbol, company name, or ISIN code.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/search")</b>: Maps HTTP GET requests to "/api/v1/market/search".</li>
     *   <li><b>@RequestParam String query</b>: Search keyword or partial string entered by the user.</li>
     * </ul>
     *
     * Internal Mechanics:
     * Cached with a 10-minute TTL per query string. Matches user input against active exchange securities.
     *
     * Real-Time Trading Scenario:
     * UI autocomplete and symbol resolution. Front-end trading terminals call this endpoint as the user types
     * to resolve fuzzy company names (e.g. "TATA MOT") to valid exchange symbols (e.g. "TATAMOTORS").
     *
     * @param query Search query text (e.g. "RELI", "INFY", "STATE BANK").
     * @return HTTP 200 OK containing matching {@link SearchResultRecord} records.
     */
    @GetMapping("/search")
    public ResponseEntity<List<SearchResultRecord>> search(@RequestParam String query) {
        log.info("REST request to search equities with query: [{}]", query);
        return ResponseEntity.ok(gatewayService.searchEquity(query));
    }
}
