package com.stockmarket.gateway.ai;

import com.stockmarket.gateway.trm.dto.TrmComprehensiveResponse;
import com.stockmarket.gateway.trm.service.TRMAlgorithmService;
import dev.langchain4j.agent.tool.P;
import dev.langchain4j.agent.tool.Tool;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

/**
 * ====================================================================================================
 *                 LANGCHAIN4J DOMAIN TOOLS: TEJA RECOVERY MODEL (TRM) EXECUTION
 * ====================================================================================================
 * 
 * Architectural Role:
 * Bridges the deterministic mathematical calculation engine (Subsystem B: Teja Recovery Model)
 * with the agentic reasoning layer (LangChain4j).
 * 
 * Why this is essential:
 * Financial LLMs struggle with complex multi-stage recursion (e.g. $B_{n+1} = 2 T_n$, $T_n = S \cdot 3^{n-1}$,
 * $99.68\%$ capital savings comparisons, and the Current Price Law for share retention).
 * By equipping the LangChain4j agent with this dedicated `@Tool`, the agent never needs to compute
 * recursive geometric progressions or OLS regressions in its head. It delegates the calculations
 * directly to the high-precision Java {@link TRMAlgorithmService}, which returns decimal-exact results.
 * 
 * Key LangChain4j Annotations Used:
 * - @Component: Exposes this class as a Spring Bean discovered by LangChain4j @AiService proxy generators.
 * - @Tool: Exposes the method `evaluateTRMModels` into the LLM's function schema catalogue.
 * - @P: Specifies parameter requirements to guide the model when generating tool arguments.
 */
@Slf4j
@Component
@RequiredArgsConstructor
public class LangChain4jTRMTools {

    private final TRMAlgorithmService trmService;

    /**
     * Tool: Complete Multi-Model TRM Algorithmic Suite & Confluence Evaluation.
     * 
     * Annotation: @Tool
     * Role: Ingests the ticker symbol, triggers the Java mathematical engine, and returns a unified
     * comparative summary of all 4 TRM models:
     * 1. TRM Classic (Recursive Cumulative Allocation, Tripling Invariant)
     * 2. TRM V1 (Percentage-Triggered Stock Equity Model)
     * 3. TRM V2 (Actual-Portfolio-Loss Adaptive Model with ~99.68% capital savings)
     * 4. TRM V3 (Asymmetric Dual-Path: Zero-cost runner retention via Current Price Law + Linear Stop Surcharge)
     * 5. 100-Period OLS Linear Regression Channel (±2.0σ Bands) and 100-Point Setup Scorecard.
     * 
     * Annotation: @P
     * Tells the LLM to supply the symbol string (e.g. "RELIANCE", "TCS", "INFY").
     * 
     * Real-Time Scenario:
     * When a user asks: "How does TRM V2 compare with V1 for Reliance if it drops 20%?", the LLM automatically
     * invokes this tool, reads the exact mathematical output from Java, and presents an authoritative
     * comparative report without mathematical calculation errors.
     * 
     * @param symbol The NSE equity ticker symbol.
     * @return Formatted analytical summary spanning all 4 TRM editions and executive action directive.
     */
    @Tool("Executes all four Teja Recovery Model (TRM) algorithms (Classic, V1, V2, V3) and 100-period OLS Linear Regression Channel confluence for a stock symbol")
    public String evaluateTRMModels(
            @P("The equity ticker symbol in uppercase, e.g. RELIANCE, TCS, INFY, SYNCOMF") String symbol) {
        log.info("[LangChain4j @Tool] Executing evaluateTRMModels for symbol: {}", symbol);
        try {
            TrmComprehensiveResponse eval = trmService.evaluateSymbol(symbol);
            return String.format("""
                    === TRM QUANTITATIVE EVALUATION FOR %s ===
                    • Last Traded Price (LTP) : ₹%s [OBSERVED]
                    • Session VWAP            : ₹%s [OBSERVED]
                    • Linear Regression Pos   : %s [CALCULATED]
                      (Mean: ₹%s | Upper +2.0σ: ₹%s | Lower -2.0σ: ₹%s)
                    • Quantitative Setup Score: %d/100 (%s) [CALCULATED]
                    
                    TEJA RECOVERY MODEL (TRM) EDITIONS:
                    1. TRM Classic (Base Model):
                       - Rounds: %d | Total Invested: ₹%s | Next Tranche: ₹%s
                       - Governing Rule: B_{n+1} = 2*T_n (Tripling Invariant Verified)
                    
                    2. TRM V1 (Percentage Drop Model):
                       - Total Invested: ₹%s | Average Cost: ₹%s | PnL: ₹%s
                    
                    3. TRM V2 (Actual Loss Adaptive Model):
                       - Total Invested: ₹%s | Average Cost: ₹%s
                       - Capital Savings vs V1: %s%% [CALCULATED]
                       - Directive: Responds strictly to actual monetary loss (B_next = 2 * ActualLoss).
                    
                    4. TRM V3 (Asymmetric Dual-Path Model):
                       - Path A (Profit Target): Target Exit: ₹%s | Liquid Cash Returned: ₹%s (100%% principal + cash profit)
                         Runner Shares Vaulted: %s shares at $0.00 effective cost basis (Current Price Law enforced).
                       - Path B (Stop Loss Exit): Stop Exit: ₹%s | Next Reversal Tranche: ₹%s
                         (Base ₹%s + Recovery Surcharge ₹%s).
                       - Universal Reset Invariant: Unconditional reset to Base S upon target fill.
                    
                    MANDATORY EXECUTIVE ACTION DIRECTIVE:
                    %s
                    """,
                    eval.getSymbol(),
                    eval.getTechnicalConfluence().getLastTradedPrice(),
                    eval.getTechnicalConfluence().getVwap(),
                    eval.getTechnicalConfluence().getChannelPosition(),
                    eval.getTechnicalConfluence().getRegressionMean(),
                    eval.getTechnicalConfluence().getUpperBand2Sigma(),
                    eval.getTechnicalConfluence().getLowerBand2Sigma(),
                    eval.getTechnicalConfluence().getSetupScore(),
                    eval.getTechnicalConfluence().getSetupClassification(),
                    eval.getTrmClassic().getMaxRounds(),
                    eval.getTrmClassic().getTotalInvested(),
                    eval.getTrmClassic().getNextInvestment(),
                    eval.getTrmV1().getTotalInvested(),
                    eval.getTrmV1().getAverageCost(),
                    eval.getTrmV1().getUnrealizedPnL(),
                    eval.getTrmV2().getTotalInvested(),
                    eval.getTrmV2().getAverageCost(),
                    eval.getTrmV2().getCapitalSavingsVsV1Percent(),
                    eval.getTrmV3().getProfitScenario().getCurrentPrice(),
                    eval.getTrmV3().getProfitScenario().getCashRealized(),
                    eval.getTrmV3().getProfitScenario().getRetainedQuantity(),
                    eval.getTrmV3().getLossScenario().getStopPrice(),
                    eval.getTrmV3().getLossScenario().getNextReversalInvestment(),
                    eval.getTrmV3().getLossScenario().getBaseReversalInvestment(),
                    eval.getTrmV3().getLossScenario().getRecoveryAmount(),
                    eval.getExecutiveDirective());
        } catch (Exception e) {
            log.error("[LangChain4j @Tool] Error executing TRM models for {}: {}", symbol, e.getMessage());
            return "Error calculating TRM models for " + symbol + ": " + e.getMessage();
        }
    }
}
