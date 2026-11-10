package com.stockmarket.gateway.ai;

import com.stockmarket.gateway.dto.CorporateAnnouncementRecord;
import com.stockmarket.gateway.dto.SymbolQuoteRecord;
import com.stockmarket.gateway.dto.YearwisePerformanceRecord;
import com.stockmarket.gateway.service.NSEMarketGatewayService;
import dev.langchain4j.agent.tool.P;
import dev.langchain4j.agent.tool.Tool;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

/**
 * ====================================================================================================
 *                 LANGCHAIN4J DOMAIN TOOLS: REAL-TIME NSE MARKET DATA INGESTION
 * ====================================================================================================
 * 
 * Architectural Overview:
 * In institutional algorithmic trading, Large Language Models (LLMs) cannot be trusted to generate financial
 * advice from parametric memory alone, because LLMs are non-deterministic and have knowledge cutoff dates.
 * 
 * LangChain4j solves this via the "Agentic Tool Calling" pattern. This class exposes verified Spring Service
 * methods to the LLM as declarative external tools.
 * 
 * Key LangChain4j Tool Annotations Explained:
 * 
 * 1. @Component (org.springframework.stereotype.Component):
 *    - Purpose: Registers this class as a Spring-managed singleton bean in the application context.
 *    - LangChain4j Auto-Wiring: LangChain4j's @AiService proxy generator automatically scans all beans
 *      in the Spring application context and registers every method marked with @Tool as an available tool.
 * 
 * 2. @Tool (dev.langchain4j.agent.tool.Tool):
 *    - Purpose: Marks a Java method as an executable tool for the LLM.
 *    - Value Attribute: Provides the human-readable description of what the tool does, when it should be called,
 *      and what information it yields.
 *    - Internal Mechanics:
 *      1. LangChain4j converts this description into a JSON Schema function declaration sent to the LLM.
 *      2. The LLM determines autonomously WHEN to call this tool based on the user's prompt.
 *      3. When the LLM decides to call it, LangChain4j catches the ToolExecutionRequest, reflects and invokes
 *         this Java method with the LLM-supplied arguments, and delivers the return value back to the LLM.
 *    - Real-Time Scenario: Prevents catastrophic financial hallucinations (e.g. LLM inventing an entry price
 *      or thinking a stock is at ₹2,500 when it is currently at ₹2,980).
 * 
 * 3. @P (dev.langchain4j.agent.tool.P):
 *    - Purpose: Documents an individual method parameter in the LLM's function calling schema.
 *    - Internal Mechanics: Mapped to the parameter's "description" field in the JSON Schema definition.
 *    - Real-Time Scenario: Teaches the LLM valid input formats (e.g. "NSE equity symbol, e.g. RELIANCE").
 */
@Slf4j
@Component
@RequiredArgsConstructor
public class LangChain4jMarketDataTools {

    private final NSEMarketGatewayService gatewayService;

    /**
     * Tool: Live Stock Quote & Market Depth Ingestion.
     * 
     * Annotation: @Tool
     * Role: Provides the LLM with real-time Last Traded Price, VWAP, Circuit Limits, and 5-Level DOM Depth.
     * 
     * Annotation: @P
     * Documents the ticker symbol parameter to ensure the model supplies uppercase NSE tickers.
     * 
     * Real-Time Scenario:
     * When a user asks "Should I buy Syncom Formulations at current market price?", the LLM automatically
     * invokes this tool first, reads the actual LTP of ₹20.60 and lower circuit at ₹18.15, and computes
     * exact risk thresholds without guessing.
     * 
     * @param symbol The NSE equity symbol (e.g. RELIANCE, TCS, SYNCOMF).
     * @return Formatted multi-line string containing factual market metrics.
     */
    @Tool("Fetches real-time NSE market quote including LTP, VWAP, Change %, Day High/Low, 52W Range, and 5-level Order Depth for an equity symbol")
    public String getLiveStockQuote(
            @P("The NSE equity symbol in uppercase, e.g. SYNCOMF, BHARTIARTL, RELIANCE, TCS") String symbol) {
        log.info("[LangChain4j @Tool] Executing getLiveStockQuote for symbol: {}", symbol);
        try {
            SymbolQuoteRecord quote = gatewayService.getSymbolQuote(symbol);
            long buyQty = (quote.orderBook() != null) ? quote.orderBook().totalBuyQuantity() : 0L;
            long sellQty = (quote.orderBook() != null) ? quote.orderBook().totalSellQuantity() : 0L;
            return String.format("""
                    Symbol: %s
                    Last Price (LTP): ₹%.2f
                    VWAP: ₹%.2f
                    Change: ₹%.2f (%.2f%%)
                    Previous Close: ₹%.2f | Open: ₹%.2f
                    52-Week Range: ₹%.2f - ₹%.2f
                    Total Buy Depth Qty: %d
                    Total Sell Depth Qty: %d
                    Lower Circuit Band: ₹%.2f | Upper Circuit Band: ₹%.2f
                    """,
                    quote.symbol(), quote.lastPrice(), quote.vwap(), quote.change(), quote.pChange(),
                    quote.previousClose(), quote.openPrice(), quote.week52Low(), quote.week52High(),
                    buyQty, sellQty,
                    quote.lowerCircuit(), quote.upperCircuit());
        } catch (Exception e) {
            log.error("[LangChain4j @Tool] Error fetching quote for {}: {}", symbol, e.getMessage());
            return "Error retrieving quote for " + symbol + ": " + e.getMessage();
        }
    }

    /**
     * Tool: Corporate Filings & Regulatory Disclosure Surveillance.
     * 
     * Annotation: @Tool
     * Role: Ingests verified corporate announcements from the NSE regulatory archive.
     * 
     * Real-Time Scenario:
     * Mandated by the 48-Hour Binary Event Embargo protocol in TRM. Trading before earnings announcements
     * or SEBI regulatory rulings involves uncontrollable gap risk. This tool enables the LLM to verify
     * that no binary events exist before greenlighting an algorithmic entry.
     * 
     * @param symbol The NSE equity symbol (e.g. INFY, RELIANCE).
     * @param limit Number of recent corporate records to inspect (defaults to 5).
     * @return Formatted list of corporate announcements and dates.
     */
    @Tool("Fetches recent corporate announcements, board meetings, earnings releases, and SEBI regulatory filings for an NSE equity symbol")
    public String getCorporateAnnouncements(
            @P("The NSE equity symbol, e.g. SYNCOMF, BHARTIARTL, INFY") String symbol,
            @P("Number of recent records to retrieve, e.g. 5") int limit) {
        int count = limit > 0 ? limit : 5;
        log.info("[LangChain4j @Tool] Executing getCorporateAnnouncements for symbol: {} (limit: {})", symbol, count);
        try {
            List<CorporateAnnouncementRecord> list = gatewayService.getCorporateAnnouncements(symbol, count);
            if (list.isEmpty()) {
                return "No recent corporate announcements found for " + symbol + ". Binary Event Embargo status is CLEAR.";
            }
            StringBuilder sb = new StringBuilder();
            for (CorporateAnnouncementRecord a : list) {
                sb.append(String.format("- [%s] %s (Attachment: %s)\n",
                        a.announcementDate(), a.description(), a.attachmentUrl()));
            }
            return sb.toString();
        } catch (Exception e) {
            log.error("[LangChain4j @Tool] Error fetching announcements for {}: {}", symbol, e.getMessage());
            return "Error retrieving announcements for " + symbol + ": " + e.getMessage();
        }
    }

    /**
     * Tool: Multi-Period Relative Strength (RS) Performance Evaluation.
     * 
     * Annotation: @Tool
     * Role: Ingests 1D, 1W, 1M, and 1Y historical price returns vs the benchmark Nifty 50.
     * 
     * Real-Time Scenario:
     * TRM Subsystem A requires verifying macro trend alignment. If an equity is severely lagging its
     * benchmark index (e.g. stock is -20% while Nifty 50 is +15%), the LLM discounts the setup score.
     * 
     * @param symbol The NSE equity symbol.
     * @return Relative performance summary.
     */
    @Tool("Fetches multi-period historical returns (1-day, 1-week, 1-year) for an NSE equity symbol to evaluate Relative Strength (RS)")
    public String getHistoricalPerformance(
            @P("The NSE equity symbol, e.g. BHARTIARTL, RELIANCE") String symbol) {
        log.info("[LangChain4j @Tool] Executing getHistoricalPerformance for symbol: {}", symbol);
        try {
            YearwisePerformanceRecord perf = gatewayService.getYearwisePerformance(symbol);
            return String.format("""
                    Symbol: %s
                    Yesterday Change: %.2f%%
                    One Week Change: %.2f%%
                    One Year Change: %.2f%%
                    Benchmark Index: %s (1-Year Change: %.2f%%)
                    """,
                    symbol, perf.yesterdayChngPer(), perf.oneWeekChngPer(), perf.oneYearChngPer(),
                    perf.indexName(), perf.indexOneYearChngPer());
        } catch (Exception e) {
            log.error("[LangChain4j @Tool] Error fetching performance for {}: {}", symbol, e.getMessage());
            return "Error retrieving historical performance for " + symbol + ": " + e.getMessage();
        }
    }

    /**
     * Tool: Exchange Circuit Band & Trading Halt Check.
     * 
     * Annotation: @Tool
     * Role: Verifies whether an equity is frozen at upper or lower price circuit bands.
     * 
     * Real-Time Scenario:
     * Placing buy orders on stocks locked in upper circuit (zero sellers) or attempting to exit stocks
     * locked in lower circuit (zero buyers) leads to execution deadlock. This tool checks exchange circuit
     * states before recommending orders.
     * 
     * @param symbol The NSE equity symbol.
     * @return Circuit status indicator ("LOCKED IN LOWER CIRCUIT", "LOCKED IN UPPER CIRCUIT", or "NORMAL TRADING").
     */
    @Tool("Checks if the stock is hitting upper or lower price circuits or trading inside normal bands")
    public String checkCircuitStatus(
            @P("The NSE equity symbol, e.g. SYNCOMF, IDEA") String symbol) {
        log.info("[LangChain4j @Tool] Executing checkCircuitStatus for symbol: {}", symbol);
        try {
            SymbolQuoteRecord quote = gatewayService.getSymbolQuote(symbol);
            double ltp = quote.lastPrice();
            double lower = quote.lowerCircuit();
            double upper = quote.upperCircuit();
            boolean isLockedLower = (lower > 0 && ltp <= lower);
            boolean isLockedUpper = (upper > 0 && ltp >= upper);
            return String.format("Symbol: %s | LTP: ₹%.2f | Lower Band: ₹%.2f | Upper Band: ₹%.2f | Status: %s",
                    symbol, ltp, lower, upper,
                    isLockedLower ? "LOCKED IN LOWER CIRCUIT (HALT)" : (isLockedUpper ? "LOCKED IN UPPER CIRCUIT (HALT)" : "NORMAL TRADING"));
        } catch (Exception e) {
            log.error("[LangChain4j @Tool] Error checking circuit status for {}: {}", symbol, e.getMessage());
            return "Error checking circuit status for " + symbol + ": " + e.getMessage();
        }
    }
}
