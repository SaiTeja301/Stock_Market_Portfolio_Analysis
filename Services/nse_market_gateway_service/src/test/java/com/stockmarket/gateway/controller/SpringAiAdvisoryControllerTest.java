package com.stockmarket.gateway.controller;

import com.stockmarket.gateway.ai.StockAdvisoryAiAgent;
import com.stockmarket.gateway.dto.AiStockAnalysisRequest;
import com.stockmarket.gateway.dto.AiStockAnalysisResponse;
import com.stockmarket.gateway.dto.ModelMetadataRecord;
import com.stockmarket.gateway.dto.TrmStructuredAdvisory;
import com.stockmarket.gateway.service.SpringAiAdvisoryService;
import dev.langchain4j.model.output.FinishReason;
import dev.langchain4j.model.output.TokenUsage;
import dev.langchain4j.service.Result;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.MockMvc;

import java.util.List;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

/**
 * Unit test suite verifying all Spring AI and LangChain4j REST endpoints in SpringAiAdvisoryController.
 */
@WebMvcTest(SpringAiAdvisoryController.class)
class SpringAiAdvisoryControllerTest {

    @Autowired
    private MockMvc mockMvc;

    @MockBean
    private SpringAiAdvisoryService aiAdvisoryService;

    @MockBean
    private StockAdvisoryAiAgent stockAdvisoryAiAgent;

    @Test
    @DisplayName("GET /api/v1/ai/models should return 200 OK with model list")
    void testGetModels() throws Exception {
        ModelMetadataRecord model = new ModelMetadataRecord(
                "qwen2.5:3b", "Qwen 2.5 3B", "Ollama", "Text & Code",
                "Advanced 3B reasoning model", "DEFAULT_QUANTITATIVE_TRM", "~2.0 GB", "Very Fast"
        );
        when(aiAdvisoryService.getAvailableModels()).thenReturn(List.of(model));

        mockMvc.perform(get("/api/v1/ai/models")
                        .accept(MediaType.APPLICATION_JSON))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$[0].modelId").value("qwen2.5:3b"))
                .andExpect(jsonPath("$[0].recommendedFor").value("DEFAULT_QUANTITATIVE_TRM"));
    }

    @Test
    @DisplayName("POST /api/v1/ai/analyze-stock should return 200 OK with Spring AI analysis")
    void testAnalyzeStock() throws Exception {
        AiStockAnalysisResponse response = new AiStockAnalysisResponse(
                "RELIANCE", "2026-09-15T10:00:00Z", "STRONG BUY", 85,
                "CLEAR", "Executive summary", List.of("Earnings growth"),
                2860.0, 3020.0, 3180.0, "Full markdown analysis", "qwen2.5:3b", "Intent matched"
        );
        when(aiAdvisoryService.analyzeStock(any())).thenReturn(response);

        mockMvc.perform(post("/api/v1/ai/analyze-stock")
                        .contentType(MediaType.APPLICATION_JSON)
                        .content("""
                                {
                                  "symbol": "RELIANCE",
                                  "analysisType": "TRM_CONFLUENCE",
                                  "userContext": "Hold 50 shares"
                                }
                                """))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.symbol").value("RELIANCE"))
                .andExpect(jsonPath("$.trmSetupScore").value(85))
                .andExpect(jsonPath("$.recommendation").value("STRONG BUY"));
    }

    @Test
    @DisplayName("POST /api/v1/ai/langchain4j/chat/{sessionId} should maintain conversation with @MemoryId")
    void testLangChain4jChat() throws Exception {
        when(stockAdvisoryAiAgent.conversationalAdvisory(eq("desk-101"), any()))
                .thenReturn("For RELIANCE, the trailing stop loss is at 2860 INR.");

        mockMvc.perform(post("/api/v1/ai/langchain4j/chat/desk-101")
                        .contentType(MediaType.APPLICATION_JSON)
                        .content("""
                                {
                                  "message": "What is the stop loss we calculated?"
                                }
                                """))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.sessionId").value("desk-101"))
                .andExpect(jsonPath("$.agentFramework").value("LangChain4j (@MemoryId + ChatMemoryProvider)"))
                .andExpect(jsonPath("$.response").value("For RELIANCE, the trailing stop loss is at 2860 INR."));
    }

    @Test
    @DisplayName("POST /api/v1/ai/langchain4j/analyze-structured should return strongly-typed Record")
    void testLangChain4jStructuredOutput() throws Exception {
        TrmStructuredAdvisory advisory = new TrmStructuredAdvisory(
                "TCS", 4215.50, 4050.00, 4420.00, 4630.00,
                88, "STRONG_BUY", "Confluence of Mean-Reversion and Bullish Momentum", true
        );
        when(stockAdvisoryAiAgent.analyzeStockStructured(eq("TCS"), any())).thenReturn(advisory);

        mockMvc.perform(post("/api/v1/ai/langchain4j/analyze-structured")
                        .contentType(MediaType.APPLICATION_JSON)
                        .content("""
                                {
                                  "symbol": "TCS",
                                  "userContext": "Evaluate fresh swing entry"
                                }
                                """))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.symbol").value("TCS"))
                .andExpect(jsonPath("$.lastTradedPrice").value(4215.50))
                .andExpect(jsonPath("$.trmStopLoss").value(4050.00))
                .andExpect(jsonPath("$.setupScore").value(88))
                .andExpect(jsonPath("$.binaryEventEmbargoClear").value(true));
    }

    @Test
    @DisplayName("GET /api/v1/ai/langchain4j/token-metadata/{symbol} should return Result token telemetry")
    void testLangChain4jTokenMetadata() throws Exception {
        TokenUsage tokenUsage = new TokenUsage(120, 85, 205);
        Result<String> mockResult = Result.<String>builder()
                .content("Rapid TRM overview: TCS is in a strong uptrend above 100 EMA.")
                .tokenUsage(tokenUsage)
                .finishReason(FinishReason.STOP)
                .build();

        when(stockAdvisoryAiAgent.analyzeWithResultMetadata("TCS")).thenReturn(mockResult);

        mockMvc.perform(get("/api/v1/ai/langchain4j/token-metadata/TCS")
                        .accept(MediaType.APPLICATION_JSON))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.symbol").value("TCS"))
                .andExpect(jsonPath("$.inputTokenCount").value(120))
                .andExpect(jsonPath("$.outputTokenCount").value(85))
                .andExpect(jsonPath("$.totalTokenCount").value(205))
                .andExpect(jsonPath("$.finishReason").value("STOP"));
    }
}
