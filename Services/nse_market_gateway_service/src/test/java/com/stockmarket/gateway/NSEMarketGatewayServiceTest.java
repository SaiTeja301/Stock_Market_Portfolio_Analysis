package com.stockmarket.gateway;

import com.stockmarket.gateway.client.NSEApiClient;
import com.stockmarket.gateway.dto.*;
import com.stockmarket.gateway.service.NSEMarketGatewayService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class NSEMarketGatewayServiceTest {

    @Mock
    private NSEApiClient apiClient;

    @InjectMocks
    private NSEMarketGatewayService gatewayService;

    private SymbolQuoteRecord mockQuote;

    @BeforeEach
    void setUp() {
        mockQuote = new SymbolQuoteRecord(
                "BHARTIARTL", "EQ", 1845.50, 15.20, 0.83, 1830.30,
                1835.00, 1845.50, 1840.10, 1655.00, 2022.00, 1900.00, 1100.00, null
        );
    }

    @Test
    @DisplayName("Should fetch normalized symbol quote data")
    void testGetSymbolQuote() {
        when(apiClient.fetchSymbolData("BHARTIARTL")).thenReturn(mockQuote);

        SymbolQuoteRecord result = gatewayService.getSymbolQuote("bhartiartl");

        assertNotNull(result);
        assertEquals("BHARTIARTL", result.symbol());
        assertEquals(1845.50, result.lastPrice());
        assertEquals(1840.10, result.vwap());
        verify(apiClient, times(1)).fetchSymbolData("BHARTIARTL");
    }

    @Test
    @DisplayName("Should fetch corporate announcements")
    void testGetCorporateAnnouncements() {
        CorporateAnnouncementRecord record = new CorporateAnnouncementRecord(
                "TMCV", "Financial Results Q1", "2026-08-10", "https://nse.com/doc.pdf",
                "Tata Motors", "INE155A01022", "Full results text", true
        );
        when(apiClient.fetchCorporateAnnouncements("TMCV", 3)).thenReturn(List.of(record));

        List<CorporateAnnouncementRecord> list = gatewayService.getCorporateAnnouncements("TMCV", 3);

        assertNotNull(list);
        assertEquals(1, list.size());
        assertEquals("Financial Results Q1", list.get(0).description());
        assertTrue(list.get(0).hasXbrl());
    }

    @Test
    @DisplayName("Should fetch intraday chart data")
    void testGetSymbolChart() {
        ChartPointRecord pt = new ChartPointRecord(1725500000000L, 1840.0, "Active", 5.0, 0.27);
        SymbolChartRecord chart = new SymbolChartRecord("BHARTIARTLEQN", "BHARTIARTL", 1840.0, List.of(pt));
        when(apiClient.fetchSymbolChartData("BHARTIARTL", "1D")).thenReturn(chart);

        SymbolChartRecord result = gatewayService.getSymbolChart("BHARTIARTL", "1D");

        assertNotNull(result);
        assertEquals(1, result.points().size());
        assertEquals(1840.0, result.points().get(0).price());
    }

    @Test
    @DisplayName("Should fetch top ten stocks snapshot")
    void testGetTopTenStocks() {
        TopTenStocksRecord record = new TopTenStocksRecord("11-Sep-2026", Collections.emptyList(), Collections.emptyList(), Collections.emptyList(), Collections.emptyList(), Collections.emptyList());
        when(apiClient.fetchTopTenStocks()).thenReturn(record);

        TopTenStocksRecord result = gatewayService.getTopTenStocks();

        assertNotNull(result);
        assertEquals("11-Sep-2026", result.timestamp());
    }

    @Test
    @DisplayName("Should fetch yearwise performance")
    void testGetYearwisePerformance() {
        YearwisePerformanceRecord perf = new YearwisePerformanceRecord(1.2, 3.4, 5.6, 8.9, 12.1, 25.4, 45.0, 80.0, 150.0, "NIFTY 50", 18.5);
        when(apiClient.fetchYearwiseData("BHARTIARTL")).thenReturn(perf);

        YearwisePerformanceRecord result = gatewayService.getYearwisePerformance("BHARTIARTL");

        assertNotNull(result);
        assertEquals(25.4, result.oneYearChngPer());
    }

    @Test
    @DisplayName("Should fetch index list")
    void testGetIndexList() {
        when(apiClient.fetchIndexList("BHARTIARTL")).thenReturn(List.of("NIFTY 50", "NIFTY 100"));

        List<String> indices = gatewayService.getIndexList("BHARTIARTL");

        assertNotNull(indices);
        assertEquals(2, indices.size());
        assertTrue(indices.contains("NIFTY 50"));
    }

    @Test
    @DisplayName("Should search equity")
    void testSearchEquity() {
        SearchResultRecord item = new SearchResultRecord("BHARTIARTL", "Bharti Airtel Limited", "EQ", "Equity", 1845.5, "/quote/BHARTIARTL");
        when(apiClient.fetchGlobalSearchEquity("BHARTI")).thenReturn(List.of(item));

        List<SearchResultRecord> list = gatewayService.searchEquity("BHARTI");

        assertNotNull(list);
        assertEquals(1, list.size());
        assertEquals("BHARTIARTL", list.get(0).symbol());
    }
}
