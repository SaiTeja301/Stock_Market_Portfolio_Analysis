package com.stockmarket.gateway.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.io.Serializable;

/**
 * ====================================================================================================
 *                 CORE MARKET DTO: CORPORATE ANNOUNCEMENT RECORD
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Immutable Data Transfer Object capturing regulatory corporate announcements, board meeting outcomes,
 * quarterly financial earnings releases, dividend declarations, and SEBI regulatory filings.
 *
 * Annotations Applied:
 * <ul>
 *   <li><b>@JsonIgnoreProperties(ignoreUnknown = true)</b>: Jackson annotation preventing deserialization
 *       errors on extraneous exchange metadata.</li>
 *   <li><b>@JsonProperty</b>: Maps exchange JSON keys to Java record parameters.</li>
 * </ul>
 *
 * Real-Time Trading Scenario:
 * Consumed by GenAI sentiment analyzers and binary risk detectors to check whether an upcoming board
 * meeting or earnings announcement poses an imminent binary event risk (gap up / gap down risk),
 * enabling TRM algorithms to freeze averaging tranches until event risk subsides.
 *
 * @param symbol           Official NSE equity ticker symbol.
 * @param description      Summary headline describing the nature of the corporate announcement.
 * @param announcementDate Broadcast timestamp string formatted by the exchange.
 * @param attachmentUrl    Direct link to the official PDF disclosure filed by company management.
 * @param companyName      Full registered corporate entity name.
 * @param isin             International Securities Identification Number (e.g., "INE002A01018").
 * @param attachmentText   Extracted plain text or OCR summary of the attached filing.
 * @param hasXbrl          Boolean indicating whether structured eXtensible Business Reporting Language
 *                         financial disclosures are available for this filing.
 */
@JsonIgnoreProperties(ignoreUnknown = true)
public record CorporateAnnouncementRecord(
        @JsonProperty("symbol") String symbol,
        @JsonProperty("description") String description,
        @JsonProperty("announcementDate") String announcementDate,
        @JsonProperty("attachmentUrl") String attachmentUrl,
        @JsonProperty("companyName") String companyName,
        @JsonProperty("isin") String isin,
        @JsonProperty("attachmentText") String attachmentText,
        @JsonProperty("hasXbrl") boolean hasXbrl
) implements Serializable {}
