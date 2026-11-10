package com.stockmarket.gateway;

import com.stockmarket.gateway.dto.*;
import com.stockmarket.gateway.service.NSEMarketGatewayService;
import com.stockmarket.gateway.service.SpringAiAdvisoryService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

@ExtendWith(MockitoExtension.class)
class SpringAiAdvisoryServiceTest {

    @Mock
    private NSEMarketGatewayService gatewayService;

    private SpringAiAdvisoryService advisoryService;

    @BeforeEach
    void setUp() {
        // Instantiate with null ChatClient to test deterministic TRM execution fallback
        advisoryService = new SpringAiAdvisoryService(null, gatewayService);
    }

    @Test
    @DisplayName("Should run deterministic TRM evaluation and generate structured response")
    void testDeterministicTrmEvaluation() {
        SymbolQuoteRecord quote = new SymbolQuoteRecord(
                "BHARTIARTL", "EQ", 1845.50, 15.20, 0.83, 1830.30,
                1835.00, 1845.50, 1840.10, 1655.00, 2022.00, 1900.00, 1100.00, null
        );
        CorporateAnnouncementRecord announcement = new CorporateAnnouncementRecord(
                "BHARTIARTL", "Board Meeting Intimation", "2026-09-01", null,
                "Bharti Airtel", "INE397D01024", "Text", false
        );

        when(gatewayService.getSymbolQuote("BHARTIARTL")).thenReturn(quote);
        when(gatewayService.getCorporateAnnouncements("BHARTIARTL", 5)).thenReturn(List.of(announcement));

        AiStockAnalysisRequest request = new AiStockAnalysisRequest("BHARTIARTL", "TRM_CONFLUENCE", null, "Invested 100 shares");
        AiStockAnalysisResponse response = advisoryService.analyzeStock(request);

        assertNotNull(response);
        assertEquals("BHARTIARTL", response.symbol());
        assertTrue(response.trmSetupScore() >= 75);
        assertEquals("TRM RE-ENTRY (CONFLUENCE VALIDATED)", response.recommendation());
        assertTrue(response.invalidationStopLoss() > 0);
        assertTrue(response.target1() > response.invalidationStopLoss());
        assertNotNull(response.fullAnalysisMarkdown());
        assertTrue(response.fullAnalysisMarkdown().contains("TRM Algorithmic Strategy Report"));
    }

    @Test
    @DisplayName("Should route models dynamically based on user explicit selection and intent")
    void testDynamicModelRouting() {
        // Scenario 1: Explicit selection
        var deepseekExplicit = advisoryService.resolveModel(new AiStockAnalysisRequest("SYNCOMF", "TRM", "deepseek", null));
        assertEquals("DeepSeek-R1:1.5b", deepseekExplicit.modelId());
        assertTrue(deepseekExplicit.routingReason().contains("EXPLICIT_USER_SELECTION"));

        var gemmaExplicit = advisoryService.resolveModel(new AiStockAnalysisRequest("SYNCOMF", "TRM", "gemma", null));
        assertEquals("gemma2:2b", gemmaExplicit.modelId());

        // Scenario 2A: Intent routing for Deep Thinking
        var thinkingIntent = advisoryService.resolveModel(new AiStockAnalysisRequest("SYNCOMF", "DEEP_THINKING", null, "Please think step by step"));
        assertEquals("DeepSeek-R1:1.5b", thinkingIntent.modelId());
        assertTrue(thinkingIntent.routingReason().contains("AUTO_ROUTED_FOR_DEEP_THINKING"));

        // Scenario 2B: Intent routing for Fast Summary
        var fastIntent = advisoryService.resolveModel(new AiStockAnalysisRequest("SYNCOMF", "FAST_SUMMARY", null, "Need quick scan"));
        assertEquals("gemma2:2b", fastIntent.modelId());
        assertTrue(fastIntent.routingReason().contains("AUTO_ROUTED_FOR_FAST_SUMMARY"));

        // Scenario 2C: Intent routing for Sentiment
        var sentimentIntent = advisoryService.resolveModel(new AiStockAnalysisRequest("SYNCOMF", "SENTIMENT", null, "Market opinion"));
        assertEquals("llama3.2:3b", sentimentIntent.modelId());
        assertTrue(sentimentIntent.routingReason().contains("AUTO_ROUTED_FOR_SENTIMENT"));

        // Scenario 2D: Default fallback to Qwen 2.5 3B
        var defaultRoute = advisoryService.resolveModel(new AiStockAnalysisRequest("SYNCOMF", "TRM_CONFLUENCE", null, null));
        assertEquals("qwen2.5:3b", defaultRoute.modelId());
        assertTrue(defaultRoute.routingReason().contains("DEFAULT_QUANTITATIVE_TRM"));
    }

    @Test
    @DisplayName("Should return available models excluding OCR models")
    void testAvailableModelsList() {
        List<ModelMetadataRecord> models = advisoryService.getAvailableModels();
        assertNotNull(models);
        assertEquals(4, models.size());
        assertTrue(models.stream().noneMatch(m -> m.modelId().toUpperCase().contains("OCR")));
        assertTrue(models.stream().anyMatch(m -> m.modelId().equals("qwen2.5:3b")));
        assertTrue(models.stream().anyMatch(m -> m.modelId().equals("DeepSeek-R1:1.5b")));
        assertTrue(models.stream().anyMatch(m -> m.modelId().equals("gemma2:2b")));
        assertTrue(models.stream().anyMatch(m -> m.modelId().equals("llama3.2:3b")));
    }

    @Test
    @DisplayName("Should detect 48-hour binary event risk from corporate actions")
    void testCheckBinaryEventRisk() {
        CorporateAnnouncementRecord riskyAnnouncement = new CorporateAnnouncementRecord(
                "BHARTIARTL", "Board Meeting to consider Unaudited Financial Results",
                "2026-09-12", null, "Bharti Airtel", "INE397D01024", "Text", false
        );
        when(gatewayService.getCorporateAnnouncements("BHARTIARTL", 5)).thenReturn(List.of(riskyAnnouncement));

        String report = advisoryService.checkBinaryEventRisk("BHARTIARTL");

        assertNotNull(report);
        assertTrue(report.contains("CAUTION"));
        assertTrue(report.contains("Mandatory TRM 48-Hour Binary Event Embargo applies"));
    }
}
