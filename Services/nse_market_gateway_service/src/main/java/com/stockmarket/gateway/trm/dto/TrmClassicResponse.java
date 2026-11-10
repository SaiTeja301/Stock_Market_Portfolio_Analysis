package com.stockmarket.gateway.trm.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.math.BigDecimal;
import java.util.List;

/**
 * ====================================================================================================
 *         TEJA RECOVERY MODEL (TRM) CLASSIC RESPONSE DTO
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Response Data Transfer Object representing the output of the TRM Classic quantitative engine.
 * Captures the sequence of bets/tranches, cumulative invested capital, required bankroll,
 * and verification of the tripling invariant.
 *
 * Mathematical Invariant Enforced:
 * <pre>
 *   B_1 = S
 *   B_{n+1} = 2 * T_n
 *   T_n = S * 3^{n-1}
 * </pre>
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@Data, @Builder, @NoArgsConstructor, @AllArgsConstructor</b>: Standard Lombok boilerplate reduction.</li>
 *   <li><b>@JsonInclude(NON_NULL)</b>: Excludes unpopulated fields.</li>
 *   <li><b>@JsonProperty</b>: Guarantees standardized JSON serialization.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Deployed in binary outcome or fixed-odds recovery simulations to determine capital survivability
 * limits and total bankroll solvency before deploying live funds.
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class TrmClassicResponse implements Serializable {

    /**
     * Formal display name of the algorithmic engine ("TRM Classic / Base TRM").
     */
    @JsonProperty("modelName")
    @Builder.Default
    private String modelName = "TRM Classic / Base TRM";

    /**
     * Initial seed investment capital S deployed at Round 1.
     */
    @JsonProperty("baseInvestment")
    private BigDecimal baseInvestment;

    /**
     * Number of rounds simulated in this progression.
     */
    @JsonProperty("maxRounds")
    private int maxRounds;

    /**
     * Chronological progression list of individual tranche allocations: [B_1, B_2, ..., B_n].
     */
    @JsonProperty("sequence")
    private List<BigDecimal> sequence;

    /**
     * Cumulative total capital invested across all simulated rounds: T_n = Sum(B_i).
     */
    @JsonProperty("totalInvested")
    private BigDecimal totalInvested;

    /**
     * Capital required for the subsequent round if the current round fails: B_{n+1} = 2 * T_n.
     */
    @JsonProperty("nextInvestment")
    private BigDecimal nextInvestment;

    /**
     * Total bankroll required to survive the next round: Bankroll = T_n + B_{n+1}.
     */
    @JsonProperty("bankrollRequired")
    private BigDecimal bankrollRequired;

    /**
     * Maximum number of consecutive losing rounds the account bankroll can survive before ruin.
     */
    @JsonProperty("maxSurvivingStreak")
    private Integer maxSurvivingStreak;

    /**
     * Verification flag indicating whether cumulative capital strictly follows the 3^(n-1) invariant.
     */
    @JsonProperty("triplingInvariantVerified")
    private boolean triplingInvariantVerified;

    /**
     * Human-readable diagnostic description of the governing mathematical recurrence.
     */
    @JsonProperty("formulaDescription")
    private String formulaDescription;

    /**
     * Granular list of {@link TrmTrancheItem} entries detailing per-round allocation mechanics.
     */
    @JsonProperty("tranches")
    private List<TrmTrancheItem> tranches;
}
