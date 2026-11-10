package com.stockmarket.gateway.ai;

import com.stockmarket.gateway.dto.CorporateAnnouncementRecord;
import com.stockmarket.gateway.dto.OrderBookRecord;
import com.stockmarket.gateway.dto.SymbolQuoteRecord;
import com.stockmarket.gateway.dto.YearwisePerformanceRecord;
import com.stockmarket.gateway.service.NSEMarketGatewayService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.when;

@ExtendWith(MockitoExtension.class)
class LangChain4jMarketDataToolsTest {

    @Mock
    private NSEMarketGatewayService gatewayService;

    private LangChain4jMarketDataTools tools;

    @BeforeEach
    void setUp() {
        tools = new LangChain4jMarketDataTools(gatewayService);
    }

    @Test
    @DisplayName("LangChain4j @Tool getLiveStockQuote should return formatted factual market metrics")
    void testGetLiveStockQuoteTool() {
        OrderBookRecord orderBook = new OrderBookRecord(
                List.of(), 50000L, 35000L
        );
        SymbolQuoteRecord quote = new SymbolQuoteRecord(
                "SYNCOMF", "EQ", 20.60, 0.45, 2.23, 20.15,
                20.20, 20.60, 20.45, 18.15, 22.15, 24.50, 12.00, orderBook
        );

        when(gatewayService.getSymbolQuote("SYNCOMF")).thenReturn(quote);

        String result = tools.getLiveStockQuote("SYNCOMF");

        assertNotNull(result);
        assertTrue(result.contains("Symbol: SYNCOMF"));
        assertTrue(result.contains("Last Price (LTP): ₹20.60"));
        assertTrue(result.contains("VWAP: ₹20.45"));
        assertTrue(result.contains("Total Buy Depth Qty: 50000"));
        assertTrue(result.contains("Total Sell Depth Qty: 35000"));
    }

    @Test
    @DisplayName("LangChain4j @Tool getCorporateAnnouncements should return formatted disclosure summary")
    void testGetCorporateAnnouncementsTool() {
        CorporateAnnouncementRecord record = new CorporateAnnouncementRecord(
                "SYNCOMF", "Financial Results Approved", "12-Sep-2026",
                "https://nsearchives.nseindia.com/ann.pdf", "Syncom Formulations",
                "INE312C01025", "Board meeting approved audited Q1 numbers", false
        );

        when(gatewayService.getCorporateAnnouncements("SYNCOMF", 5)).thenReturn(List.of(record));

        String result = tools.getCorporateAnnouncements("SYNCOMF", 5);

        assertNotNull(result);
        assertTrue(result.contains("Financial Results Approved"));
        assertTrue(result.contains("12-Sep-2026"));
    }

    @Test
    @DisplayName("LangChain4j @Tool getHistoricalPerformance should return multi-period returns")
    void testGetHistoricalPerformanceTool() {
        YearwisePerformanceRecord perf = new YearwisePerformanceRecord(
                1.5, 3.2, 5.8, 12.0, 18.5, 45.0, 60.0, 85.0, 150.0, "NIFTY 50", 22.0
        );

        when(gatewayService.getYearwisePerformance("BHARTIARTL")).thenReturn(perf);

        String result = tools.getHistoricalPerformance("BHARTIARTL");

        assertNotNull(result);
        assertTrue(result.contains("Yesterday Change: 1.50%"));
        assertTrue(result.contains("One Week Change: 3.20%"));
        assertTrue(result.contains("One Year Change: 45.00%"));
        assertTrue(result.contains("Benchmark Index: NIFTY 50"));
    }

    @Test
    @DisplayName("LangChain4j @Tool checkCircuitStatus should detect normal trading vs locked circuits")
    void testCheckCircuitStatusTool() {
        SymbolQuoteRecord quote = new SymbolQuoteRecord(
                "SYNCOMF", "EQ", 20.60, 0.45, 2.23, 20.15,
                20.20, 20.60, 20.45, 18.15, 22.15, 24.50, 12.00, null
        );

        when(gatewayService.getSymbolQuote("SYNCOMF")).thenReturn(quote);

        String result = tools.checkCircuitStatus("SYNCOMF");

        assertNotNull(result);
        assertTrue(result.contains("NORMAL TRADING"));
    }
}
