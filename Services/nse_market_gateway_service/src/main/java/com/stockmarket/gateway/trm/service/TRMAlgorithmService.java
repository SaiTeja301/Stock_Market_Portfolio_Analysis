package com.stockmarket.gateway.trm.service;

import com.stockmarket.gateway.trm.dto.*;

/**
 * ====================================================================================================
 *         TEJA RECOVERY MODEL (TRM) ALGORITHMIC SERVICE CONTRACT
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Formal interface defining the operational contracts for the four Teja Recovery Model (TRM) algorithms,
 * multi-model comparative benchmarking, and real-time exchange quote synchronization.
 *
 * The Four Algorithmic Paradigms:
 * <ol>
 *   <li><b>TRM Classic (Recursive Cumulative Capital Allocation)</b>:
 *       Designed for binary and fixed-return trading instruments. Solves for next-bet sizing where
 *       $B_{n+1} = 2 \cdot T_n$, ensuring net profit upon recovery equals initial base capital $S$.</li>
 *   <li><b>TRM V1 (Percentage-Triggered Stock Equity Model)</b>:
 *       Systematic scale-in averaging on percentage drawdowns ($P_n = P_0 (1-L)^n$). Computes weighted
 *       average purchase prices and breakeven exit thresholds.</li>
 *   <li><b>TRM V2 (Adaptive Portfolio Loss Model)</b>:
 *       Capital-preservation engine calculating recovery tranche capital based strictly on open
 *       unrealized mark-to-market loss ($L_{actual} = T_n - Q_n \cdot P_{curr}$). Conserves ~99.68%
 *       more capital than naive martingale schemes.</li>
 *   <li><b>TRM V3 (Asymmetric Dual-Path Engine)</b>:
 *       Advanced trade execution model enforcing the <i>Current Price Law</i> ($Q_{retain} = R / P_{curr}$)
 *       to retain permanent zero-cost equity shares ("runners") upon reaching profit targets, while
 *       calculating systematic compensatory allocations if stopped out.</li>
 * </ol>
 *
 * Real-Time Trading Scenario:
 * Decouples the presentation layer (REST controllers, AI decision agents, Spring AI Function Callbacks,
 * and LangChain4j tools) from underlying mathematical compute engines and upstream exchange data fetchers.
 */
public interface TRMAlgorithmService {

    /**
     * Computes the TRM Classic recursive capital allocation sequence.
     *
     * <p><b>Mathematical Governing Recurrence:</b></p>
     * <pre>
     *   B_1 = S
     *   B_{n+1} = 2 * T_n
     *   T_n = S * 3^{n-1}
     * </pre>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Employed when calculating bankroll solvency requirements for fixed-odds or binary outcome recovery
     * sequences across 1 to N stages.
     *
     * @param request Input parameters containing base capital, max rounds, and optional bankroll.
     * @return {@link TrmClassicResponse} containing tranche progression, cumulative totals, and safety warnings.
     */
    TrmClassicResponse calculateClassic(TrmCalculationRequest request);

    /**
     * Computes the TRM V1 percentage-triggered equity averaging model.
     *
     * <p><b>Mathematical Governing Equation:</b></p>
     * <pre>
     *   P_n = P_0 * (1 - L)^n
     *   C_n = C_0 * M^n
     *   Q_n = C_n / P_n
     *   WAP_n = Sum(C_i) / Sum(Q_i)
     * </pre>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Used by long-only equity accumulation bots scaling into high-quality growth stocks during sustained
     * sector corrections.
     *
     * @param request Input parameters containing base capital, entry price, drop percentage, and multiplier.
     * @return {@link TrmV1Response} containing per-round tranche prices, share quantities, and weighted average price.
     */
    TrmV1Response calculateV1(TrmCalculationRequest request);

    /**
     * Computes the TRM V2 adaptive actual-portfolio-loss recovery model.
     *
     * <p><b>Mathematical Governing Equation:</b></p>
     * <pre>
     *   Drawdown = TotalCost - (TotalShares * CurrentPrice)
     *   NextCapital = TargetRecoveryCapital(Drawdown)
     * </pre>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Deployed by institutional risk desks to average positions without exponential capital explosion,
     * maintaining strict single-stock bankroll risk caps (<= 15%).
     *
     * @param request Input parameters containing base capital, entry price, drop percentage, and bankroll.
     * @return {@link TrmV2Response} containing round metrics and quantitative capital savings vs Classic.
     */
    TrmV2Response calculateV2(TrmCalculationRequest request);

    /**
     * Computes the TRM V3 asymmetric dual-path engine (Profit runner retention + Stop-loss recovery).
     *
     * <p><b>Mathematical Governing Equation (Current Price Law):</b></p>
     * <pre>
     *   Q_retain = NetProfit / P_current
     *   Q_liquidate = TotalShares - Q_retain
     *   PrincipalHarvested = Q_liquidate * P_current >= InitialCapital
     * </pre>
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Executed by breakout momentum desks to lock in profit, recover all initial seed capital, and hold
     * completely risk-free runner shares indefinitely for multi-year compounding.
     *
     * @param request Input parameters specifying profit target percentage, stop loss percentage, and target return.
     * @return {@link TrmV3Response} detailing both profit liquidation/runner retention and stop-loss recovery paths.
     */
    TrmV3Response calculateV3(TrmCalculationRequest request);

    /**
     * Executes all four TRM algorithmic models concurrently against a single calculation request,
     * computing technical regression channel confluence when price history is present.
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Strategy benchmarking engines and backtesting frameworks evaluating which TRM model variant is
     * mathematically optimal for a specific equity profile.
     *
     * @param request Consolidated calculation parameters.
     * @return {@link TrmComprehensiveResponse} encapsulating results from all 4 models and confluence data.
     */
    TrmComprehensiveResponse evaluateAll(TrmCalculationRequest request);

    /**
     * Fetches live market quotes, order book depth, and historical intraday chart points for the
     * specified NSE ticker symbol, and executes a full comprehensive TRM multi-model evaluation.
     *
     * <p><b>Real-Time Trading Scenario:</b></p>
     * Autonomous trade execution systems or user-facing dashboard widgets requesting a complete,
     * actionable TRM strategy evaluation for a live stock (e.g. "RELIANCE", "INFY", "TATAMOTORS").
     *
     * @param symbol NSE equity ticker symbol.
     * @return {@link TrmComprehensiveResponse} populated with live prices, regression channels, and TRM models.
     */
    TrmComprehensiveResponse evaluateSymbol(String symbol);
}
