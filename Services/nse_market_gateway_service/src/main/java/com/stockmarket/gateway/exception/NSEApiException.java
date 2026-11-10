package com.stockmarket.gateway.exception;

/**
 * ====================================================================================================
 *              CUSTOM DOMAIN EXCEPTION: NSE API CLIENT ERROR REPRESENTATION
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Represents unchecked domain-specific exceptions originating from interactions with the National Stock
 * Exchange (NSE) upstream endpoints or internal API processing pipelines. Encapsulates upstream HTTP
 * status codes, network protocol errors, Akamai anti-bot perimeter challenges (HTTP 401/403), rate limit
 * exhaustions, and JSON deserialization failures.
 *
 * Role in System Design:
 * Acts as the unified domain exception boundary between low-level network clients (Reactor Netty / WebClient)
 * and high-level service/controller layers. Prevents raw Netty or Jackson exceptions from propagating
 * unhandled, enabling {@link GlobalExceptionHandler} to intercept them and produce standardized RFC-7807
 * {@link org.springframework.http.ProblemDetail} responses with the exact HTTP status code.
 *
 * Real-Time Trading Scenario:
 * During volatile market openings (09:15 IST), upstream exchange endpoints may throttle traffic (HTTP 429)
 * or return transient 502/504 Bad Gateway errors. Raising {@code NSEApiException} triggers Resilience4j
 * circuit breaker fallback metrics or graceful cache degradation, preventing trading desk timeouts.
 */
public class NSEApiException extends RuntimeException {

    /**
     * The HTTP status code associated with the upstream or downstream failure (e.g., 400, 404, 429, 500, 502).
     */
    private final int statusCode;

    /**
     * Constructs a new {@code NSEApiException} with a default HTTP 500 Internal Server Error status.
     *
     * @param message Human-readable diagnostic description of the error.
     */
    public NSEApiException(String message) {
        super(message);
        this.statusCode = 500;
    }

    /**
     * Constructs a new {@code NSEApiException} with a descriptive message and an explicit HTTP status code.
     *
     * @param message    Human-readable diagnostic description of the failure.
     * @param statusCode Upstream or synthesized HTTP status code (e.g., 404 for invalid symbol, 429 for rate limit).
     */
    public NSEApiException(String message, int statusCode) {
        super(message);
        this.statusCode = statusCode;
    }

    /**
     * Constructs a new {@code NSEApiException} with a descriptive message, the underlying root cause,
     * and an explicit HTTP status code.
     *
     * @param message    Human-readable diagnostic description of the failure.
     * @param cause      The underlying throwable root cause (e.g., WebClientResponseException, JsonProcessingException).
     * @param statusCode Upstream or synthesized HTTP status code.
     */
    public NSEApiException(String message, Throwable cause, int statusCode) {
        super(message, cause);
        this.statusCode = statusCode;
    }

    /**
     * Retrieves the HTTP status code mapped to this exception.
     *
     * @return Integer representing standard HTTP status code (e.g. 200, 400, 404, 429, 500, 502).
     */
    public int getStatusCode() {
        return statusCode;
    }
}
