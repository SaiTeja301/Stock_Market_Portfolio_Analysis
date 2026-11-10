package com.stockmarket.gateway.trm.controller;

import com.stockmarket.gateway.trm.dto.*;
import com.stockmarket.gateway.trm.service.TRMAlgorithmService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

/**
 * ====================================================================================================
 *         TEJA RECOVERY MODEL (TRM) ALGORITHMIC SUITE REST CONTROLLER
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Exposes an enterprise-grade RESTful API orchestrating the 4 proprietary Teja Recovery Model (TRM)
 * quantitative capital allocation and drawdown recovery engines:
 * <ol>
 *   <li><b>TRM Classic (Base Recursive Model)</b>: Recursive cumulative capital allocation governed by the
 *       tripling invariant (B_{n+1} = 2 * T_n, where T_n = S * 3^{n-1}). Designed for binary outcomes.</li>
 *   <li><b>TRM V1 (Percentage-Triggered Equity Model)</b>: Systematic dip averaging triggered at fixed
 *       geometric drawdowns (P_n = P_0 * (1 - L)^n) across predefined tranches.</li>
 *   <li><b>TRM V2 (Adaptive Portfolio Loss Model)</b>: Precision capital sizing calculated strictly against
 *       actual open mark-to-market drawdown rather than cumulative theoretical exposure, yielding
 *       ~99.68% capital savings compared to Classic Martingale approaches.</li>
 *   <li><b>TRM V3 (Asymmetric Dual-Path Engine)</b>: Asymmetric execution delivering permanent zero-cost
 *       equity runners via the <i>Current Price Law</i> (Q_retain = R / P_current) upon profit targets,
 *       and automated recovery surcharges upon stop-loss invalidations.</li>
 * </ol>
 *
 * Supports dual invocation paradigms:
 * <ul>
 *   <li><b>Manual Input Simulation</b>: Via JSON payload supplying custom capital, entry prices, and tranches.</li>
 *   <li><b>Live Market Data Evaluation</b>: Via exchange symbol URI paths fetching real-time Level-1 quotes
 *       and 100-period Ordinary Least Squares (OLS) regression channels from NSE.</li>
 * </ul>
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@RestController</b>: Specialized Spring component stereotype combining {@code @Controller}
 *       and {@code @ResponseBody}. Directs Spring's DispatcherServlet to serialize all handler return
 *       types directly into the HTTP response body via Jackson ObjectMapper.</li>
 *   <li><b>@RequestMapping("/api/v1/trm")</b>: Sets the global URI path prefix for all TRM endpoints,
 *       enforcing strict versioning and microservice route isolation.</li>
 *   <li><b>@RequiredArgsConstructor</b>: Lombok annotation synthesizing a constructor initializing the
 *       {@code final TRMAlgorithmService} dependency, guaranteeing immutability and testability.</li>
 *   <li><b>@Slf4j</b>: Lombok annotation injecting an asynchronous SLF4J logger instance for operational
 *       telemetry, audit trails, and execution diagnostics.</li>
 * </ul>
 */
@Slf4j
@RestController
@RequestMapping("/api/v1/trm")
@RequiredArgsConstructor
public class TRMAlgorithmController {

    private final TRMAlgorithmService trmService;

    /**
     * Executes the TRM Classic algorithmic engine using custom simulation parameters.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@PostMapping("/classic")</b>: Routes HTTP POST requests to this method for complex calculation requests.</li>
     *   <li><b>@RequestBody(required = false)</b>: Deserializes the incoming JSON body into {@link TrmCalculationRequest}.
     *       If omitted, defaults to fallback parameters ($100 base, 5 rounds).</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Quantitative analysts backtesting binary options or fixed-return recovery sequences to evaluate bankroll
     * requirements and geometric risk escalations before deploying capital.
     *
     * @param request Input parameters containing baseCapital, maxRounds, and bankroll.
     * @return {@link ResponseEntity} wrapping {@link TrmClassicResponse} with detailed tranche tables.
     */
    @PostMapping("/classic")
    public ResponseEntity<TrmClassicResponse> calculateClassic(@RequestBody(required = false) TrmCalculationRequest request) {
        log.info("Calculating TRM Classic model: {}", request);
        return ResponseEntity.ok(trmService.calculateClassic(request != null ? request : new TrmCalculationRequest()));
    }

    /**
     * Executes the TRM Classic engine dynamically seeded with real-time market data for an NSE symbol.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@GetMapping("/classic/{symbol}")</b>: Binds HTTP GET requests for quick single-stock lookups.</li>
     *   <li><b>@PathVariable</b>: Extracts the target ticker symbol (e.g., "RELIANCE", "TCS") from the URI path.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Traders needing an instantaneous calculation of Classic recovery tranches anchored to the live Last
     * Traded Price (LTP) of a specified NSE stock.
     *
     * @param symbol NSE equity ticker symbol.
     * @return {@link ResponseEntity} wrapping {@link TrmClassicResponse}.
     */
    @GetMapping("/classic/{symbol}")
    public ResponseEntity<TrmClassicResponse> calculateClassicForSymbol(@PathVariable String symbol) {
        log.info("Calculating TRM Classic model for symbol: {}", symbol);
        TrmCalculationRequest request = TrmCalculationRequest.builder().symbol(symbol).build();
        return ResponseEntity.ok(trmService.calculateClassic(request));
    }

    /**
     * Executes the TRM V1 percentage-triggered equity averaging model with custom parameters.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@PostMapping("/v1")</b>: Maps HTTP POST requests to the V1 averaging engine.</li>
     *   <li><b>@RequestBody(required = false)</b>: Deserializes JSON payload into {@link TrmCalculationRequest}.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Swing traders planning fixed-percentage dip accumulation (e.g., 20% drops with 2.0x capital multiplier)
     * to determine target breakeven prices, weighted average prices, and shares per tranche.
     *
     * @param request Calculation request containing baseCapital, entryPrice, dropPercentage, and multiplier.
     * @return {@link ResponseEntity} wrapping {@link TrmV1Response}.
     */
    @PostMapping("/v1")
    public ResponseEntity<TrmV1Response> calculateV1(@RequestBody(required = false) TrmCalculationRequest request) {
        log.info("Calculating TRM V1 model: {}", request);
        return ResponseEntity.ok(trmService.calculateV1(request != null ? request : new TrmCalculationRequest()));
    }

    /**
     * Executes the TRM V1 engine dynamically seeded with real-time market data for an NSE symbol.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@GetMapping("/v1/{symbol}")</b>: Maps GET requests to retrieve live V1 levels.</li>
     *   <li><b>@PathVariable</b>: Extracts the target ticker symbol from the URI.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Long-term positional investors establishing scale-in limit orders on high-conviction blue-chip stocks
     * based on current market valuation.
     *
     * @param symbol NSE equity ticker symbol.
     * @return {@link ResponseEntity} wrapping {@link TrmV1Response}.
     */
    @GetMapping("/v1/{symbol}")
    public ResponseEntity<TrmV1Response> calculateV1ForSymbol(@PathVariable String symbol) {
        log.info("Calculating TRM V1 model for symbol: {}", symbol);
        TrmCalculationRequest request = TrmCalculationRequest.builder().symbol(symbol).build();
        return ResponseEntity.ok(trmService.calculateV1(request));
    }

    /**
     * Executes the TRM V2 actual-loss adaptive model using manual simulation parameters.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@PostMapping("/v2")</b>: Maps HTTP POST requests to the V2 adaptive engine.</li>
     *   <li><b>@RequestBody(required = false)</b>: Deserializes incoming JSON or applies safe defaults.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Institutional risk managers seeking conservative, capital-efficient drawdown averaging that scales
     * tranche size strictly to cover accumulated unrealized loss rather than doubling total invested capital.
     *
     * @param request Calculation request containing baseCapital, entryPrice, dropPercentage, and bankroll.
     * @return {@link ResponseEntity} wrapping {@link TrmV2Response} showcasing capital savings.
     */
    @PostMapping("/v2")
    public ResponseEntity<TrmV2Response> calculateV2(@RequestBody(required = false) TrmCalculationRequest request) {
        log.info("Calculating TRM V2 model: {}", request);
        return ResponseEntity.ok(trmService.calculateV2(request != null ? request : new TrmCalculationRequest()));
    }

    /**
     * Executes the TRM V2 model dynamically seeded with live market quotes for an NSE symbol.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@GetMapping("/v2/{symbol}")</b>: Maps GET requests for symbol-specific V2 metrics.</li>
     *   <li><b>@PathVariable</b>: Target equity ticker symbol.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Discretionary desk traders evaluating real-time dip protection levels and capital conservation metrics
     * before entering an accumulation phase.
     *
     * @param symbol NSE equity ticker symbol.
     * @return {@link ResponseEntity} wrapping {@link TrmV2Response}.
     */
    @GetMapping("/v2/{symbol}")
    public ResponseEntity<TrmV2Response> calculateV2ForSymbol(@PathVariable String symbol) {
        log.info("Calculating TRM V2 model for symbol: {}", symbol);
        TrmCalculationRequest request = TrmCalculationRequest.builder().symbol(symbol).build();
        return ResponseEntity.ok(trmService.calculateV2(request));
    }

    /**
     * Executes the TRM V3 asymmetric dual-path engine with custom simulation parameters.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@PostMapping("/v3")</b>: Maps HTTP POST requests to the V3 dual-path engine.</li>
     *   <li><b>@RequestBody(required = false)</b>: Deserializes incoming parameters or falls back to defaults.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Advanced algorithmic trading desks modeling asymmetric risk-reward: automatically harvesting principal
     * and locking in "free shares" (runners) via the Current Price Law, while calculating compensatory
     * position sizes if stopped out.
     *
     * @param request Calculation request specifying profitTargetPct, stopLossPct, targetReturn, and entryPrice.
     * @return {@link ResponseEntity} wrapping {@link TrmV3Response} with separate Profit and Loss scenarios.
     */
    @PostMapping("/v3")
    public ResponseEntity<TrmV3Response> calculateV3(@RequestBody(required = false) TrmCalculationRequest request) {
        log.info("Calculating TRM V3 model: {}", request);
        return ResponseEntity.ok(trmService.calculateV3(request != null ? request : new TrmCalculationRequest()));
    }

    /**
     * Executes the TRM V3 asymmetric engine dynamically seeded with real-time NSE market quotes.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@GetMapping("/v3/{symbol}")</b>: Maps GET requests to retrieve live V3 dual-path execution levels.</li>
     *   <li><b>@PathVariable</b>: Target equity ticker symbol.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Active momentum traders entering breakout setups who need immediate calculation of exact runner
     * retention quantities and secondary stop-loss recovery requirements.
     *
     * @param symbol NSE equity ticker symbol.
     * @return {@link ResponseEntity} wrapping {@link TrmV3Response}.
     */
    @GetMapping("/v3/{symbol}")
    public ResponseEntity<TrmV3Response> calculateV3ForSymbol(@PathVariable String symbol) {
        log.info("Calculating TRM V3 model for symbol: {}", symbol);
        TrmCalculationRequest request = TrmCalculationRequest.builder().symbol(symbol).build();
        return ResponseEntity.ok(trmService.calculateV3(request));
    }

    /**
     * Executes a comprehensive multi-model comparative evaluation across all 4 TRM algorithms concurrently.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@PostMapping("/evaluate-all")</b>: Maps POST requests to the composite evaluation pipeline.</li>
     *   <li><b>@RequestBody(required = false)</b>: Input parameters applied uniformly across all 4 models.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Portfolio managers conducting side-by-side risk, capital drawdown, and equity yield comparisons
     * across Classic, V1, V2, and V3 strategies to select the optimal model for prevailing market conditions.
     *
     * @param request Simulation parameters.
     * @return {@link ResponseEntity} wrapping {@link TrmComprehensiveResponse} containing all 4 model outputs.
     */
    @PostMapping("/evaluate-all")
    public ResponseEntity<TrmComprehensiveResponse> evaluateAll(@RequestBody(required = false) TrmCalculationRequest request) {
        log.info("Executing comprehensive evaluation across all 4 TRM models: {}", request);
        return ResponseEntity.ok(trmService.evaluateAll(request != null ? request : new TrmCalculationRequest()));
    }

    /**
     * Executes live comprehensive evaluation for a specified NSE ticker symbol, integrating real-time
     * quote, order book depth, and 100-period OLS Linear Regression Channel technical confluence.
     *
     * <p><b>Annotations Applied:</b></p>
     * <ul>
     *   <li><b>@GetMapping("/evaluate/{symbol}")</b>: Binds GET requests for full multi-engine live evaluation.</li>
     *   <li><b>@PathVariable</b>: Target equity ticker symbol.</li>
     * </ul>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Autonomous trading systems and AI decision agents requesting a unified technical and quantitative
     * recovery dossier for a live stock prior to dispatching algorithmic orders to the broker exchange.
     *
     * @param symbol NSE equity ticker symbol.
     * @return {@link ResponseEntity} wrapping {@link TrmComprehensiveResponse} with live technical confluence.
     */
    @GetMapping("/evaluate/{symbol}")
    public ResponseEntity<TrmComprehensiveResponse> evaluateSymbol(@PathVariable String symbol) {
        log.info("Executing live comprehensive TRM evaluation for symbol: {}", symbol);
        return ResponseEntity.ok(trmService.evaluateSymbol(symbol));
    }
}
