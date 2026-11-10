package com.stockmarket.research.trm;

import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

/**
 * Standalone Reference Implementation for all four Teja Recovery Model (TRM) Algorithms:
 * 1. TRM Classic (Recursive Cumulative Capital Allocation, B_{n+1} = 2*T_n)
 * 2. TRM V1 (Percentage-Triggered Stock Equity Model, P_n = P_0 * (1 - L)^n)
 * 3. TRM V2 (Actual-Portfolio-Loss Adaptive Model, B_{next} = m * ActualLoss)
 * 4. TRM V3 (Asymmetric Dual-Path Model: Runner Retention via Current Price Law + Surcharged Reversal)
 */
public class TRMStandaloneEngine {

    private static final MathContext MC = MathContext.DECIMAL128;
    private static final int CURRENCY_SCALE = 2;
    private static final int SHARE_SCALE = 6;
    private static final RoundingMode ROUNDING = RoundingMode.HALF_UP;

    // --- Data Records ---
    public record ClassicResult(List<BigDecimal> sequence, BigDecimal totalInvested, BigDecimal nextInvestment, BigDecimal bankrollRequired) {}
    
    public record TrancheRow(int round, BigDecimal price, BigDecimal investment, BigDecimal shares, BigDecimal totalInvested, BigDecimal totalShares, BigDecimal avgCost, BigDecimal marketValue, BigDecimal pnl, BigDecimal actualLoss) {}

    public record V3ProfitResult(BigDecimal entryPrice, BigDecimal currentPrice, BigDecimal quantity, BigDecimal initialInvestment, BigDecimal profitAmount, BigDecimal retainedValue, BigDecimal retainedQuantity, BigDecimal quantityToSell, BigDecimal cashRealized, BigDecimal netCashProfit) {}

    public record V3LossResult(BigDecimal entryPrice, BigDecimal stopPrice, BigDecimal quantity, BigDecimal initialInvestment, BigDecimal lossAmount, BigDecimal preservedCash, BigDecimal recoveryAmount, BigDecimal baseReversal, BigDecimal nextReversal) {}

    // --- 1. TRM Classic ---
    public static ClassicResult runClassic(BigDecimal base, int maxRounds) {
        BigDecimal bet = base;
        BigDecimal total = BigDecimal.ZERO;
        List<BigDecimal> seq = new ArrayList<>();
        for (int i = 1; i <= maxRounds; i++) {
            seq.add(bet.setScale(CURRENCY_SCALE, ROUNDING));
            total = total.add(bet, MC);
            bet = total.multiply(BigDecimal.valueOf(2), MC);
        }
        return new ClassicResult(seq, total.setScale(CURRENCY_SCALE, ROUNDING), bet.setScale(CURRENCY_SCALE, ROUNDING), total.add(bet, MC).setScale(CURRENCY_SCALE, ROUNDING));
    }

    // --- 2. TRM V1 ---
    public static List<TrancheRow> runV1(BigDecimal initialInvestment, BigDecimal initialPrice, BigDecimal dropPct, int maxRounds, BigDecimal multiplier) {
        List<TrancheRow> list = new ArrayList<>();
        BigDecimal currentPrice = initialPrice;
        BigDecimal totalInvested = initialInvestment;
        BigDecimal totalShares = initialInvestment.divide(initialPrice, MC);
        BigDecimal remaining = BigDecimal.ONE.subtract(dropPct, MC);

        list.add(new TrancheRow(1, currentPrice.setScale(CURRENCY_SCALE, ROUNDING), initialInvestment.setScale(CURRENCY_SCALE, ROUNDING), totalShares.setScale(SHARE_SCALE, ROUNDING), totalInvested.setScale(CURRENCY_SCALE, ROUNDING), totalShares.setScale(SHARE_SCALE, ROUNDING), currentPrice.setScale(CURRENCY_SCALE, ROUNDING), initialInvestment.setScale(CURRENCY_SCALE, ROUNDING), BigDecimal.ZERO.setScale(CURRENCY_SCALE, ROUNDING), BigDecimal.ZERO.setScale(CURRENCY_SCALE, ROUNDING)));

        for (int r = 2; r <= maxRounds + 1; r++) {
            currentPrice = currentPrice.multiply(remaining, MC);
            BigDecimal nextInvest = totalInvested.multiply(multiplier, MC);
            BigDecimal sharesBought = nextInvest.divide(currentPrice, MC);
            totalInvested = totalInvested.add(nextInvest, MC);
            totalShares = totalShares.add(sharesBought, MC);
            BigDecimal avgCost = totalInvested.divide(totalShares, MC);
            BigDecimal mv = totalShares.multiply(currentPrice, MC);
            BigDecimal pnl = mv.subtract(totalInvested, MC);

            list.add(new TrancheRow(r, currentPrice.setScale(CURRENCY_SCALE, ROUNDING), nextInvest.setScale(CURRENCY_SCALE, ROUNDING), sharesBought.setScale(SHARE_SCALE, ROUNDING), totalInvested.setScale(CURRENCY_SCALE, ROUNDING), totalShares.setScale(SHARE_SCALE, ROUNDING), avgCost.setScale(CURRENCY_SCALE, ROUNDING), mv.setScale(CURRENCY_SCALE, ROUNDING), pnl.setScale(CURRENCY_SCALE, ROUNDING), BigDecimal.ZERO));
        }
        return list;
    }

    // --- 3. TRM V2 ---
    public static List<TrancheRow> runV2(BigDecimal initialInvestment, BigDecimal initialPrice, BigDecimal dropPct, int maxRounds, BigDecimal multiplier) {
        List<TrancheRow> list = new ArrayList<>();
        BigDecimal currentPrice = initialPrice;
        BigDecimal totalInvested = initialInvestment;
        BigDecimal totalShares = initialInvestment.divide(initialPrice, MC);
        BigDecimal remaining = BigDecimal.ONE.subtract(dropPct, MC);

        list.add(new TrancheRow(1, currentPrice.setScale(CURRENCY_SCALE, ROUNDING), initialInvestment.setScale(CURRENCY_SCALE, ROUNDING), totalShares.setScale(SHARE_SCALE, ROUNDING), totalInvested.setScale(CURRENCY_SCALE, ROUNDING), totalShares.setScale(SHARE_SCALE, ROUNDING), currentPrice.setScale(CURRENCY_SCALE, ROUNDING), initialInvestment.setScale(CURRENCY_SCALE, ROUNDING), BigDecimal.ZERO, BigDecimal.ZERO));

        for (int r = 2; r <= maxRounds + 1; r++) {
            currentPrice = currentPrice.multiply(remaining, MC);
            BigDecimal mv = totalShares.multiply(currentPrice, MC);
            BigDecimal actualLoss = totalInvested.subtract(mv, MC).max(BigDecimal.ZERO);
            BigDecimal nextInvest = actualLoss.multiply(multiplier, MC);
            BigDecimal sharesBought = nextInvest.divide(currentPrice, MC);

            totalInvested = totalInvested.add(nextInvest, MC);
            totalShares = totalShares.add(sharesBought, MC);
            BigDecimal avgCost = totalInvested.divide(totalShares, MC);
            mv = totalShares.multiply(currentPrice, MC);
            BigDecimal pnl = mv.subtract(totalInvested, MC);

            list.add(new TrancheRow(r, currentPrice.setScale(CURRENCY_SCALE, ROUNDING), nextInvest.setScale(CURRENCY_SCALE, ROUNDING), sharesBought.setScale(SHARE_SCALE, ROUNDING), totalInvested.setScale(CURRENCY_SCALE, ROUNDING), totalShares.setScale(SHARE_SCALE, ROUNDING), avgCost.setScale(CURRENCY_SCALE, ROUNDING), mv.setScale(CURRENCY_SCALE, ROUNDING), pnl.setScale(CURRENCY_SCALE, ROUNDING), actualLoss.setScale(CURRENCY_SCALE, ROUNDING)));
        }
        return list;
    }

    // --- 4. TRM V3 ---
    public static V3ProfitResult runV3Profit(BigDecimal entryPrice, BigDecimal quantity, BigDecimal profitPct, BigDecimal retentionFactor) {
        BigDecimal initialInvest = entryPrice.multiply(quantity, MC);
        BigDecimal currentPrice = entryPrice.multiply(BigDecimal.ONE.add(profitPct, MC), MC);
        BigDecimal profitAmount = initialInvest.multiply(profitPct, MC);
        BigDecimal retainedValue = profitAmount.multiply(retentionFactor, MC);
        // Current Price Law:
        BigDecimal retainedQuantity = retainedValue.divide(currentPrice, MC);
        BigDecimal quantityToSell = quantity.subtract(retainedQuantity, MC);
        BigDecimal cashRealized = quantityToSell.multiply(currentPrice, MC);
        BigDecimal netCashProfit = profitAmount.multiply(BigDecimal.ONE.subtract(retentionFactor, MC), MC);

        return new V3ProfitResult(
                entryPrice.setScale(CURRENCY_SCALE, ROUNDING),
                currentPrice.setScale(CURRENCY_SCALE, ROUNDING),
                quantity.setScale(SHARE_SCALE, ROUNDING),
                initialInvest.setScale(CURRENCY_SCALE, ROUNDING),
                profitAmount.setScale(CURRENCY_SCALE, ROUNDING),
                retainedValue.setScale(CURRENCY_SCALE, ROUNDING),
                retainedQuantity.setScale(SHARE_SCALE, ROUNDING),
                quantityToSell.setScale(SHARE_SCALE, ROUNDING),
                cashRealized.setScale(CURRENCY_SCALE, ROUNDING),
                netCashProfit.setScale(CURRENCY_SCALE, ROUNDING)
        );
    }

    public static V3LossResult runV3Loss(BigDecimal entryPrice, BigDecimal quantity, BigDecimal lossPct, BigDecimal recoveryFactor, BigDecimal baseReversal) {
        BigDecimal initialInvest = entryPrice.multiply(quantity, MC);
        BigDecimal stopPrice = entryPrice.multiply(BigDecimal.ONE.subtract(lossPct, MC), MC);
        BigDecimal lossAmount = initialInvest.multiply(lossPct, MC);
        BigDecimal preservedCash = quantity.multiply(stopPrice, MC);
        BigDecimal recoveryAmount = lossAmount.multiply(recoveryFactor, MC);
        BigDecimal nextReversal = baseReversal.add(recoveryAmount, MC);

        return new V3LossResult(
                entryPrice.setScale(CURRENCY_SCALE, ROUNDING),
                stopPrice.setScale(CURRENCY_SCALE, ROUNDING),
                quantity.setScale(SHARE_SCALE, ROUNDING),
                initialInvest.setScale(CURRENCY_SCALE, ROUNDING),
                lossAmount.setScale(CURRENCY_SCALE, ROUNDING),
                preservedCash.setScale(CURRENCY_SCALE, ROUNDING),
                recoveryAmount.setScale(CURRENCY_SCALE, ROUNDING),
                baseReversal.setScale(CURRENCY_SCALE, ROUNDING),
                nextReversal.setScale(CURRENCY_SCALE, ROUNDING)
        );
    }

    public static void main(String[] args) {
        System.out.println("================================================================================");
        System.out.println("TEJA RECOVERY MODEL (TRM) - ALL 4 EDITIONS BENCHMARK EXECUTION");
        System.out.println("================================================================================\n");

        // 1. TRM Classic
        ClassicResult classic = runClassic(new BigDecimal("5.00"), 10);
        System.out.printf("1. TRM Classic (S=5, N=10): Total Invested: ₹%s | Next Tranche: ₹%s | Bankroll Required: ₹%s%n",
                classic.totalInvested(), classic.nextInvestment(), classic.bankrollRequired());

        // 2. TRM V1 vs V2 Comparison
        List<TrancheRow> v1 = runV1(new BigDecimal("100.00"), new BigDecimal("18.00"), new BigDecimal("0.20"), 10, new BigDecimal("2.0"));
        List<TrancheRow> v2 = runV2(new BigDecimal("100.00"), new BigDecimal("18.00"), new BigDecimal("0.20"), 10, new BigDecimal("2.0"));

        BigDecimal v1Total = v1.get(10).totalInvested();
        BigDecimal v2Total = v2.get(10).totalInvested();
        BigDecimal savings = v1Total.subtract(v2Total).divide(v1Total, MC).multiply(new BigDecimal("100")).setScale(2, ROUNDING);

        System.out.printf("2. TRM V1 (10 drops of 20%%): Total Invested = ₹%s | Avg Cost = ₹%s%n", v1Total, v1.get(10).avgCost());
        System.out.printf("3. TRM V2 (10 drops of 20%%): Total Invested = ₹%s | Avg Cost = ₹%s%n", v2Total, v2.get(10).avgCost());
        System.out.printf("   >>> Capital Savings of TRM V2 vs TRM V1: %s%%%n%n", savings);

        // 4. TRM V3 Dual-Path
        V3ProfitResult p = runV3Profit(new BigDecimal("200.00"), new BigDecimal("10.00"), new BigDecimal("0.10"), new BigDecimal("0.20"));
        System.out.printf("4. TRM V3 Profit Path (AAPL +10%%, 20%% retention):%n");
        System.out.printf("   Target: $%s | Cash Returned: $%s (100%% principal + $%s cash profit)%n", p.currentPrice(), p.cashRealized(), p.netCashProfit());
        System.out.printf("   Vaulted Runner: %s shares at $0.00 cost basis!%n", p.retainedQuantity());

        V3LossResult l = runV3Loss(new BigDecimal("200.00"), new BigDecimal("10.00"), new BigDecimal("0.05"), new BigDecimal("0.20"), new BigDecimal("1000.00"));
        System.out.printf("   TRM V3 Loss Path (AAPL -5%% stop, 20%% recovery factor):%n");
        System.out.printf("   Stop Exit: $%s | Preserved Cash: $%s | Surcharge: $%s | Next Trade Allocation: $%s%n",
                l.stopPrice(), l.preservedCash(), l.recoveryAmount(), l.nextReversal());
        System.out.println("================================================================================");
    }
}
