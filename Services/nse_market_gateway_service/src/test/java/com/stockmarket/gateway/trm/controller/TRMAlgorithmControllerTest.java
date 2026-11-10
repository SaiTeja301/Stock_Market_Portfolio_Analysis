package com.stockmarket.gateway.trm.controller;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.stockmarket.gateway.trm.dto.*;
import com.stockmarket.gateway.trm.service.TRMAlgorithmService;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.MockMvc;

import java.math.BigDecimal;
import java.util.List;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

@WebMvcTest(TRMAlgorithmController.class)
class TRMAlgorithmControllerTest {

    @Autowired
    private MockMvc mockMvc;

    @Autowired
    private ObjectMapper objectMapper;

    @MockBean
    private TRMAlgorithmService trmService;

    @Test
    @DisplayName("POST /api/v1/trm/classic returns TRM Classic response")
    void testCalculateClassicPost() throws Exception {
        TrmClassicResponse mockResp = TrmClassicResponse.builder()
                .modelName("TRM Classic / Base TRM")
                .baseInvestment(new BigDecimal("100.00"))
                .maxRounds(5)
                .totalInvested(new BigDecimal("8100.00"))
                .nextInvestment(new BigDecimal("16200.00"))
                .sequence(List.of(new BigDecimal("100.00"), new BigDecimal("200.00")))
                .triplingInvariantVerified(true)
                .build();

        when(trmService.calculateClassic(any(TrmCalculationRequest.class))).thenReturn(mockResp);

        TrmCalculationRequest req = TrmCalculationRequest.builder()
                .baseCapital(new BigDecimal("100.00"))
                .maxRounds(5)
                .build();

        mockMvc.perform(post("/api/v1/trm/classic")
                        .contentType(MediaType.APPLICATION_JSON)
                        .content(objectMapper.writeValueAsString(req)))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.modelName").value("TRM Classic / Base TRM"))
                .andExpect(jsonPath("$.totalInvested").value(8100.00))
                .andExpect(jsonPath("$.triplingInvariantVerified").value(true));
    }

    @Test
    @DisplayName("GET /api/v1/trm/classic/{symbol} returns TRM Classic for stock")
    void testCalculateClassicGetSymbol() throws Exception {
        TrmClassicResponse mockResp = TrmClassicResponse.builder()
                .modelName("TRM Classic / Base TRM")
                .baseInvestment(new BigDecimal("28000.00"))
                .maxRounds(5)
                .build();

        when(trmService.calculateClassic(any(TrmCalculationRequest.class))).thenReturn(mockResp);

        mockMvc.perform(get("/api/v1/trm/classic/RELIANCE"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.baseInvestment").value(28000.00));
    }

    @Test
    @DisplayName("POST /api/v1/trm/v1 returns TRM V1 percentage-triggered response")
    void testCalculateV1Post() throws Exception {
        TrmV1Response mockResp = TrmV1Response.builder()
                .modelName("TRM V1 - Percentage-Triggered Model")
                .initialPrice(new BigDecimal("18.00"))
                .totalInvested(new BigDecimal("5904900.00"))
                .averageCost(new BigDecimal("2.13"))
                .build();

        when(trmService.calculateV1(any(TrmCalculationRequest.class))).thenReturn(mockResp);

        mockMvc.perform(post("/api/v1/trm/v1")
                        .contentType(MediaType.APPLICATION_JSON)
                        .content("{}"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.modelName").value("TRM V1 - Percentage-Triggered Model"))
                .andExpect(jsonPath("$.averageCost").value(2.13));
    }

    @Test
    @DisplayName("POST /api/v1/trm/v2 returns TRM V2 actual-loss response with capital savings")
    void testCalculateV2Post() throws Exception {
        TrmV2Response mockResp = TrmV2Response.builder()
                .modelName("TRM V2 - Actual-Loss Adaptive Model")
                .initialPrice(new BigDecimal("18.00"))
                .totalInvested(new BigDecimal("18712.11"))
                .averageCost(new BigDecimal("2.45"))
                .capitalSavingsVsV1Percent(new BigDecimal("99.68"))
                .build();

        when(trmService.calculateV2(any(TrmCalculationRequest.class))).thenReturn(mockResp);

        mockMvc.perform(post("/api/v1/trm/v2")
                        .contentType(MediaType.APPLICATION_JSON)
                        .content("{}"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.modelName").value("TRM V2 - Actual-Loss Adaptive Model"))
                .andExpect(jsonPath("$.totalInvested").value(18712.11))
                .andExpect(jsonPath("$.capitalSavingsVsV1Percent").value(99.68));
    }

    @Test
    @DisplayName("POST /api/v1/trm/v3 returns TRM V3 dual-path response (Profit & Loss)")
    void testCalculateV3Post() throws Exception {
        TrmV3ProfitResult profitResult = TrmV3ProfitResult.builder()
                .entryPrice(new BigDecimal("200.00"))
                .currentPrice(new BigDecimal("220.00"))
                .profitAmount(new BigDecimal("200.00"))
                .retainedQuantity(new BigDecimal("0.181818"))
                .cashRealized(new BigDecimal("2160.00"))
                .effectiveCostBasis(BigDecimal.ZERO)
                .reconciliationReconciled(true)
                .build();

        TrmV3LossResult lossResult = TrmV3LossResult.builder()
                .entryPrice(new BigDecimal("200.00"))
                .stopPrice(new BigDecimal("190.00"))
                .lossAmount(new BigDecimal("100.00"))
                .recoveryAmount(new BigDecimal("20.00"))
                .nextReversalInvestment(new BigDecimal("1020.00"))
                .bankrollCapPassed(true)
                .build();

        TrmV3Response mockResp = TrmV3Response.builder()
                .modelName("TRM V3 - Asymmetric Dual-Path Model")
                .profitScenario(profitResult)
                .lossScenario(lossResult)
                .currentPriceLawVerified(true)
                .build();

        when(trmService.calculateV3(any(TrmCalculationRequest.class))).thenReturn(mockResp);

        mockMvc.perform(post("/api/v1/trm/v3")
                        .contentType(MediaType.APPLICATION_JSON)
                        .content("{}"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.modelName").value("TRM V3 - Asymmetric Dual-Path Model"))
                .andExpect(jsonPath("$.profitScenario.retainedQuantity").value(0.181818))
                .andExpect(jsonPath("$.profitScenario.cashRealized").value(2160.00))
                .andExpect(jsonPath("$.lossScenario.nextReversalInvestment").value(1020.00))
                .andExpect(jsonPath("$.currentPriceLawVerified").value(true));
    }

    @Test
    @DisplayName("GET /api/v1/trm/evaluate/{symbol} returns comprehensive multi-model response")
    void testEvaluateSymbolGet() throws Exception {
        TrmTechnicalConfluence confluence = TrmTechnicalConfluence.builder()
                .symbol("RELIANCE")
                .lastTradedPrice(new BigDecimal("2950.00"))
                .channelPosition("LOWER_BAND_DEMAND_ZONE (-2.0σ High Probability Entry)")
                .setupScore(88)
                .setupClassification("STRONG (High Conviction Setup)")
                .build();

        TrmComprehensiveResponse mockResp = TrmComprehensiveResponse.builder()
                .symbol("RELIANCE")
                .timestamp("2026-09-15T00:00:00Z")
                .technicalConfluence(confluence)
                .trmClassic(TrmClassicResponse.builder().totalInvested(new BigDecimal("28000.00")).build())
                .trmV1(TrmV1Response.builder().totalInvested(new BigDecimal("84000.00")).build())
                .trmV2(TrmV2Response.builder().totalInvested(new BigDecimal("33600.00")).capitalSavingsVsV1Percent(new BigDecimal("99.68")).build())
                .trmV3(TrmV3Response.builder().build())
                .comparisonSummary("TRM Comparison Summary")
                .executiveDirective("ACTION: HIGH CONVICTION VALUE ENTRY on RELIANCE")
                .build();

        when(trmService.evaluateSymbol(eq("RELIANCE"))).thenReturn(mockResp);

        mockMvc.perform(get("/api/v1/trm/evaluate/RELIANCE"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.symbol").value("RELIANCE"))
                .andExpect(jsonPath("$.technicalConfluence.setupScore").value(88))
                .andExpect(jsonPath("$.executiveDirective").value("ACTION: HIGH CONVICTION VALUE ENTRY on RELIANCE"));
    }
}
