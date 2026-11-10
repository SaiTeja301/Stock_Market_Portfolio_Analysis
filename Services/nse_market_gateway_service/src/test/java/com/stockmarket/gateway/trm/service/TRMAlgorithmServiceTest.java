package com.stockmarket.gateway.trm.service;

import com.stockmarket.gateway.dto.OrderBookRecord;
import com.stockmarket.gateway.dto.SymbolQuoteRecord;
import com.stockmarket.gateway.service.NSEMarketGatewayService;
import com.stockmarket.gateway.trm.dto.*;
import com.stockmarket.gateway.trm.engine.TrmMathematicalEngine;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import java.math.BigDecimal;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.when;

@ExtendWith(MockitoExtension.class)
class TRMAlgorithmServiceTest {

    @Mock
    private NSEMarketGatewayService marketGatewayService;

    private TRMAlgorithmServiceImpl trmService;

    @BeforeEach
    void setUp() {
        trmService = new TRMAlgorithmServiceImpl(marketGatewayService);
    }

    @Test
    @DisplayName("TRM Classic: Verifies Induction Sequence, Bankroll Formula and Tripling Invariant (S=5, N=10)")
    void testTRMClassic_InductionSequence() {
        TrmCalculationRequest request = TrmCalculationRequest.builder()
                .baseCapital(new BigDecimal("5.00"))
                .maxRounds(10)
                .bankroll(new BigDecimal("405.00"))
                .build();

        TrmClassicResponse response = trmService.calculateClassic(request);

        assertNotNull(response);
        assertEquals(10, response.getSequence().size());
        assertEquals(new BigDecimal("5.00"), response.getSequence().get(0));
        assertEquals(new BigDecimal("10.00"), response.getSequence().get(1));
        assertEquals(new BigDecimal("30.00"), response.getSequence().get(2));
        assertEquals(new BigDecimal("90.00"), response.getSequence().get(3));
        assertEquals(new BigDecimal("270.00"), response.getSequence().get(4));

        // Round 10 total: T_10 = 5 * 3^9 = 98415.00
        assertEquals(new BigDecimal("98415.00"), response.getTotalInvested());

        // Next investment: B_11 = 2 * T_10 = 196830.00
        assertEquals(new BigDecimal("196830.00"), response.getNextInvestment());

        // Bankroll required: M = 5 * 3^10 = 295245.00
        assertEquals(new BigDecimal("295245.00"), response.getBankrollRequired());

        // Max surviving streak for bankroll=405: floor(log3(405/5)) = floor(log3(81)) = 4
        assertEquals(4, response.getMaxSurvivingStreak());
        assertTrue(response.isTriplingInvariantVerified());
    }

    @Test
    @DisplayName("TRM V1: Verifies 10-Drop Exponential Progression (P0=18, S=100, L=20%)")
    void testTRMV1_TenDropProgression() {
        TrmCalculationRequest request = TrmCalculationRequest.builder()
                .baseCapital(new BigDecimal("100.00"))
                .entryPrice(new BigDecimal("18.00"))
                .dropPercentage(new BigDecimal("0.20"))
                .maxRounds(10)
                .multiplier(new BigDecimal("2.0"))
                .build();

        TrmV1Response response = trmService.calculateV1(request);

        assertNotNull(response);
        assertEquals(11, response.getSteps().size());

        // Step 1
        TrmTrancheItem step1 = response.getSteps().get(0);
        assertEquals(new BigDecimal("18.00"), step1.getStockPrice());
        assertEquals(new BigDecimal("100.00"), step1.getTotalInvested());
        assertEquals(new BigDecimal("18.00"), step1.getAverageCost());

        // Step 2: Price = 14.40, Tranche = 200, Total = 300, AC ~ 15.43
        TrmTrancheItem step2 = response.getSteps().get(1);
        assertEquals(new BigDecimal("14.40"), step2.getStockPrice());
        assertEquals(new BigDecimal("200.00"), step2.getTrancheInvestment());
        assertEquals(new BigDecimal("300.00"), step2.getTotalInvested());
        assertEquals(new BigDecimal("15.43"), step2.getAverageCost());

        // Step 11: Total Invested = 5,904,900.00, Average Cost = 2.13
        TrmTrancheItem step11 = response.getSteps().get(10);
        assertEquals(new BigDecimal("1.93"), step11.getStockPrice());
        assertEquals(new BigDecimal("5904900.00"), step11.getTotalInvested());
        assertEquals(new BigDecimal("2.13"), step11.getAverageCost());
    }

    @Test
    @DisplayName("TRM V2: Verifies 99.68% Capital Savings vs V1 over 10 consecutive drops")
    void testTRMV2_CapitalSavingsBenchmark() {
        TrmCalculationRequest request = TrmCalculationRequest.builder()
                .baseCapital(new BigDecimal("100.00"))
                .entryPrice(new BigDecimal("18.00"))
                .dropPercentage(new BigDecimal("0.20"))
                .maxRounds(10)
                .multiplier(new BigDecimal("2.0"))
                .build();

        TrmV2Response response = trmService.calculateV2(request);

        assertNotNull(response);
        assertEquals(11, response.getSteps().size());

        // Step 1: Initial purchase 100 at 18
        TrmTrancheItem step1 = response.getSteps().get(0);
        assertEquals(new BigDecimal("18.00"), step1.getStockPrice());
        assertEquals(new BigDecimal("100.00"), step1.getTotalInvested());

        // Step 2: Price falls to 14.40 -> Actual Loss = 20.00 -> Tranche = 40.00 -> Total = 140.00 -> AC = 16.80
        TrmTrancheItem step2 = response.getSteps().get(1);
        assertEquals(new BigDecimal("14.40"), step2.getStockPrice());
        assertEquals(new BigDecimal("40.00"), step2.getTrancheInvestment());
        assertEquals(new BigDecimal("140.00"), step2.getTotalInvested());
        assertEquals(new BigDecimal("16.80"), step2.getAverageCost());

        // Step 11: Total Invested = 18,712.11 vs V1's 5,904,900.00 (99.68% savings)
        TrmTrancheItem step11 = response.getSteps().get(10);
        assertEquals(new BigDecimal("1.93"), step11.getStockPrice());
        assertEquals(new BigDecimal("18712.11"), step11.getTotalInvested());
        assertEquals(new BigDecimal("2.45"), step11.getAverageCost());
        assertEquals(new BigDecimal("7959.94"), step11.getTrancheInvestment());

        // Capital savings percentage check
        BigDecimal savings = response.getCapitalSavingsVsV1Percent();
        assertTrue(savings.compareTo(new BigDecimal("99.60")) >= 0, "Capital savings must be ~99.68%");
    }

    @Test
    @DisplayName("TRM V2: Large-Cap Case Study (Reliance Industries @ 2800 -> 2520)")
    void testTRMV2_RelianceCaseStudy() {
        BigDecimal entryPrice = new BigDecimal("2800.00");
        BigDecimal totalShares = new BigDecimal("10.00");
        BigDecimal totalInvested = new BigDecimal("28000.00");
        BigDecimal currentPrice = new BigDecimal("2520.00"); // -10% drop
        BigDecimal multiplier = new BigDecimal("2.0");

        TrmV2Response v2 = TrmMathematicalEngine.computeV2SingleStep(
                currentPrice, totalShares, totalInvested, multiplier, new BigDecimal("300000.00")
        );

        // Market Value: 10 * 2520 = 25,200.00
        assertEquals(new BigDecimal("25200.00"), v2.getMarketValue());
        // Actual Loss: 28,000 - 25,200 = 2,800.00
        assertEquals(new BigDecimal("2800.00"), v2.getActualLoss());
        // Next Investment: 2 * 2800 = 5,600.00
        assertEquals(new BigDecimal("5600.00"), v2.getNextInvestment());
        // Additional Shares: 5600 / 2520 = 2.222222
        assertEquals(new BigDecimal("2.222222"), v2.getAdditionalShares());
        // New Total Invested: 33,600.00
        assertEquals(new BigDecimal("33600.00"), v2.getTotalInvested());
        // New Average Cost: 33600 / 12.222222 = 2,749.09
        assertEquals(new BigDecimal("2749.09"), v2.getAverageCost());
        assertFalse(v2.isBankrollCapBreached());
    }

    @Test
    @DisplayName("TRM V3: Apple Swing Trade - Profit Scenario (The Current Price Law Proof)")
    void testTRMV3_AppleProfitPath() {
        BigDecimal entryPrice = new BigDecimal("200.00");
        BigDecimal quantity = new BigDecimal("10.00");
        BigDecimal profitPct = new BigDecimal("0.10"); // +10%
        BigDecimal retentionFactor = new BigDecimal("0.20"); // 20% retention

        TrmV3ProfitResult profit = TrmMathematicalEngine.computeV3Profit(
                entryPrice, quantity, profitPct, retentionFactor
        );

        assertEquals(new BigDecimal("200.00"), profit.getEntryPrice());
        assertEquals(new BigDecimal("220.00"), profit.getCurrentPrice());
        assertEquals(new BigDecimal("2000.00"), profit.getInitialInvestment());
        assertEquals(new BigDecimal("200.00"), profit.getProfitAmount());
        assertEquals(new BigDecimal("40.00"), profit.getRetainedProfitValue());

        // CURRENT PRICE LAW: Q_retain = 40.00 / 220.00 = 0.181818 shares
        assertEquals(new BigDecimal("0.181818"), profit.getRetainedQuantity());
        // Shares to sell: 10 - 0.181818 = 9.818182 shares
        assertEquals(new BigDecimal("9.818182"), profit.getQuantityToSell());
        // Cash realized: 9.818182 * 220.00 = $2,160.00 (100% initial $2,000 + $160 cash profit)
        assertEquals(new BigDecimal("2160.00"), profit.getCashRealized());
        assertEquals(new BigDecimal("160.00"), profit.getNetCashProfit());
        // Runner equity value: 0.181818 * 220.00 = $40.00
        assertEquals(new BigDecimal("40.00"), profit.getRunnerEquityValue());
        assertEquals(BigDecimal.ZERO.setScale(2), profit.getEffectiveCostBasis());
        assertTrue(profit.isReconciliationReconciled());
    }

    @Test
    @DisplayName("TRM V3: Apple Swing Trade - Loss Scenario (Linear Surcharge Calibration)")
    void testTRMV3_AppleLossPath() {
        BigDecimal entryPrice = new BigDecimal("200.00");
        BigDecimal quantity = new BigDecimal("10.00");
        BigDecimal lossPct = new BigDecimal("0.05"); // -5% stop loss
        BigDecimal recoveryFactor = new BigDecimal("0.20");
        BigDecimal baseReversal = new BigDecimal("1000.00");
        BigDecimal bankroll = new BigDecimal("50000.00");

        TrmV3LossResult loss = TrmMathematicalEngine.computeV3Loss(
                entryPrice, quantity, lossPct, recoveryFactor, baseReversal, bankroll
        );

        assertEquals(new BigDecimal("200.00"), loss.getEntryPrice());
        assertEquals(new BigDecimal("190.00"), loss.getStopPrice());
        assertEquals(new BigDecimal("2000.00"), loss.getInitialInvestment());
        assertEquals(new BigDecimal("100.00"), loss.getLossAmount());
        assertEquals(new BigDecimal("1900.00"), loss.getPreservedCash());
        // Recovery Surcharge: 100 * 0.20 = 20.00
        assertEquals(new BigDecimal("20.00"), loss.getRecoveryAmount());
        // Next trade: 1000 + 20 = 1020.00
        assertEquals(new BigDecimal("1020.00"), loss.getNextReversalInvestment());
        assertTrue(loss.isBankrollCapPassed());
    }

    @Test
    @DisplayName("Linear Regression Channel: Validates OLS Slope, Intercept, Mean, and Envelopes")
    void testLinearRegressionChannel() {
        // Line: Price = 100 + 2 * t (t = 1..5)
        List<Double> prices = List.of(102.0, 104.0, 106.0, 108.0, 110.0);
        TrmTechnicalConfluence confluence = TrmMathematicalEngine.computeRegressionChannel(
                "TCS", prices, 110.0, 106.0, 120.0, 95.0, 88.0, 132.0
        );

        assertNotNull(confluence);
        assertEquals("TCS", confluence.getSymbol());
        assertEquals(new BigDecimal("2.0000"), confluence.getRegressionSlope());
        assertEquals(new BigDecimal("100.00"), confluence.getRegressionIntercept());
        assertEquals(new BigDecimal("110.00"), confluence.getRegressionMean());
        assertEquals(new BigDecimal("0.00"), confluence.getStandardError());
        assertEquals(new BigDecimal("110.00"), confluence.getUpperBand2Sigma());
        assertEquals(new BigDecimal("110.00"), confluence.getLowerBand2Sigma());
    }

    @Test
    @DisplayName("Comprehensive Evaluation: Integrates live quote data into all 4 models")
    void testEvaluateSymbolWithMockQuote() {
        OrderBookRecord orderBook = new OrderBookRecord(List.of(), 50000L, 30000L);
        SymbolQuoteRecord mockQuote = new SymbolQuoteRecord(
                "RELIANCE", "EQ", 2950.00, 25.00, 0.85, 2925.00,
                2930.00, 2950.00, 2942.00, 2632.50, 3217.50, 3024.90, 2220.30, orderBook
        );

        when(marketGatewayService.getSymbolQuote("RELIANCE")).thenReturn(mockQuote);

        TrmComprehensiveResponse response = trmService.evaluateSymbol("RELIANCE");

        assertNotNull(response);
        assertEquals("RELIANCE", response.getSymbol());
        assertNotNull(response.getTechnicalConfluence());
        assertEquals(new BigDecimal("2950.00"), response.getTechnicalConfluence().getLastTradedPrice());
        assertNotNull(response.getTrmClassic());
        assertNotNull(response.getTrmV1());
        assertNotNull(response.getTrmV2());
        assertNotNull(response.getTrmV3());
        assertNotNull(response.getComparisonSummary());
        assertNotNull(response.getExecutiveDirective());
    }
}
