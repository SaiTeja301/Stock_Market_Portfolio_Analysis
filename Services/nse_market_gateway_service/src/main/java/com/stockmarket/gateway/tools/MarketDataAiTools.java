package com.stockmarket.gateway.tools;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.stockmarket.gateway.dto.CorporateAnnouncementRecord;
import com.stockmarket.gateway.dto.SymbolChartRecord;
import com.stockmarket.gateway.dto.SymbolQuoteRecord;
import com.stockmarket.gateway.dto.YearwisePerformanceRecord;
import com.stockmarket.gateway.service.NSEMarketGatewayService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Description;

import java.util.List;
import java.util.function.Function;

/**
 * ====================================================================================================
 *                 SPRING AI FUNCTION CALLBACKS & TOOL REGISTRATION BEANS
 * ====================================================================================================
 * 
 * Architectural Overview:
 * Spring AI utilizes standard Java {@link java.util.function.Function} beans to implement Function Calling
 * (Tool Calling). Rather than inventing proprietary annotations for tools, Spring AI leverages Spring's
 * existing `@Bean` and `@Description` infrastructure.
 * 
 * Key Spring AI Tool Calling Concepts Explained:
 * 
 * 1. @Bean Function<Request, Response>:
 *    - Purpose: Exposes a Java function as a callable tool for LLMs in the Spring container.
 *    - Internal Mechanics:
 *      1. Spring AI inspects the generic input type `Request` (e.g. `StockSymbolRequest`) and uses Jackson
 *         to generate the OpenAPI JSON Schema describing the function's arguments.
 *      2. When ChatClient is instructed to use tools via `.tools("getLiveStockQuoteTool", ...)`,
 *         Spring AI sends these function definitions to Ollama / OpenAI.
 *      3. When the LLM issues a function-call tool request, Spring AI deserializes the arguments into `Request`,
 *         executes the `apply(request)` lambda, serializes the `Response` back to JSON, and sends it to the LLM.
 * 
 * 2. @Description (org.springframework.context.annotation.Description):
 *    - Purpose: Provides the description of the tool to the LLM.
 *    - Real-Time Scenario: Teaches the LLM what data this tool returns and when it should decide to invoke it.
 * 
 * Comparison with LangChain4j:
 * - Spring AI: Uses standard `@Bean Function<Req, Resp>` with `@Description`.
 * - LangChain4j: Uses `@Tool` on ordinary bean methods with `@P` on method parameters.
 */
@Slf4j
@Configuration
@RequiredArgsConstructor
public class MarketDataAiTools {

    private final NSEMarketGatewayService gatewayService;

    // --- Strongly-Typed Request Payloads for Function Calling Schema Generation ---

    /**
     * Request schema for ticker symbol lookups.
     * Jackson @JsonProperty annotations dictate required fields and schema property names.
     */
    public record StockSymbolRequest(
            @JsonProperty(required = true, value = "symbol") String symbol
    ) {}

    /**
     * Request schema for historical chart time-series lookups.
     */
    public record ChartDataRequest(
            @JsonProperty(required = true, value = "symbol") String symbol,
            @JsonProperty(value = "days") String days
    ) {}

    /**
     * Request schema for corporate disclosures and filings.
     */
    public record AnnouncementsRequest(
            @JsonProperty(required = true, value = "symbol") String symbol,
            @JsonProperty(value = "noOfRecords") Integer noOfRecords
    ) {}

    /**
     * Tool: Real-Time NSE Market Quote Function Callback.
     * 
     * Spring AI Mechanics:
     * Exposes live LTP, VWAP, circuit limits, and DOM depth to the Spring AI ChatClient.
     * 
     * @return Java Function mapping StockSymbolRequest -> SymbolQuoteRecord.
     */
    @Bean
    @Description("Fetches real-time market quote, LTP, VWAP, circuit limits, and 5-level order book depth for an NSE equity symbol")
    public Function<StockSymbolRequest, SymbolQuoteRecord> getLiveStockQuoteTool() {
        return request -> {
            log.info("[Spring AI Tool] getLiveStockQuoteTool invoked for symbol: [{}]", request.symbol());
            return gatewayService.getSymbolQuote(request.symbol());
        };
    }

    /**
     * Tool: Corporate Announcement Surveillance Function Callback.
     * 
     * Spring AI Mechanics:
     * Ingests regulatory filings from the NSE backend to verify binary event risks.
     * 
     * @return Java Function mapping AnnouncementsRequest -> List<CorporateAnnouncementRecord>.
     */
    @Bean
    @Description("Fetches latest corporate announcements, board meetings, earnings releases, and SEBI regulatory filings for an NSE equity symbol")
    public Function<AnnouncementsRequest, List<CorporateAnnouncementRecord>> getCorporateAnnouncementsTool() {
        return request -> {
            int count = (request.noOfRecords() != null && request.noOfRecords() > 0) ? request.noOfRecords() : 5;
            log.info("[Spring AI Tool] getCorporateAnnouncementsTool invoked for symbol: [{}] (limit: {})", request.symbol(), count);
            return gatewayService.getCorporateAnnouncements(request.symbol(), count);
        };
    }

    /**
     * Tool: Relative Strength Historical Performance Function Callback.
     * 
     * Spring AI Mechanics:
     * Ingests 1D, 1W, 1M, and 1Y performance vs benchmark index for macro trend filtering.
     * 
     * @return Java Function mapping StockSymbolRequest -> YearwisePerformanceRecord.
     */
    @Bean
    @Description("Fetches historical multi-period relative strength and performance percentages (1D, 1W, 1M, 1Y, 3Y, 5Y) for an NSE symbol")
    public Function<StockSymbolRequest, YearwisePerformanceRecord> getHistoricalPerformanceTool() {
        return request -> {
            log.info("[Spring AI Tool] getHistoricalPerformanceTool invoked for symbol: [{}]", request.symbol());
            return gatewayService.getYearwisePerformance(request.symbol());
        };
    }

    /**
     * Tool: Intraday Timestamped Price Series Function Callback.
     * 
     * Spring AI Mechanics:
     * Ingests time-series tick arrays for candlestick pattern and regression channel reconstruction.
     * 
     * @return Java Function mapping ChartDataRequest -> SymbolChartRecord.
     */
    @Bean
    @Description("Fetches intraday timestamped price series for candlestick reconstruction for an NSE symbol")
    public Function<ChartDataRequest, SymbolChartRecord> getIntradayChartTool() {
        return request -> {
            String days = request.days() != null ? request.days() : "1D";
            log.info("[Spring AI Tool] getIntradayChartTool invoked for symbol: [{}] (days: {})", request.symbol(), days);
            return gatewayService.getSymbolChart(request.symbol(), days);
        };
    }
}
