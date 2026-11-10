package com.stockmarket.gateway.exception;

import io.github.resilience4j.circuitbreaker.CallNotPermittedException;
import io.github.resilience4j.ratelimiter.RequestNotPermitted;
import org.springframework.http.HttpStatus;
import org.springframework.http.ProblemDetail;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

import java.net.URI;
import java.time.Instant;

/**
 * ====================================================================================================
 *         GLOBAL REST EXCEPTION HANDLER: RFC-7807 PROBLEM DETAILS SPECIFICATION
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Centralized exception interceptor and error translator for the entire microservice.
 * Translates low-level runtime errors, upstream exchange failures, and Resilience4j circuit breaker
 * trips into standardized RFC-7807 {@link ProblemDetail} JSON payloads.
 *
 * Why RFC-7807 in Institutional Trading:
 * Algorithmic order execution systems, automated hedging daemons, and trader frontends require
 * predictable, machine-readable JSON error schemas. Uncaught runtime exceptions returning generic
 * 500 HTML error pages break downstream deserializers. This handler ensures every single error
 * response contains title, status, human-readable detail, error type URI, and an ISO-8601 timestamp.
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@RestControllerAdvice</b>: Combines Spring's {@code @ControllerAdvice} and {@code @ResponseBody}.
 *       Intercepts exceptions thrown by any handler method in {@code @RestController} classes across
 *       the entire application and serializes the return value into the HTTP response body.</li>
 * </ul>
 */
@RestControllerAdvice
public class GlobalExceptionHandler {

    /**
     * Handles upstream exchange communication failures and HTTP error statuses from NSE.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@ExceptionHandler(NSEApiException.class)</b>: Directs Spring MVC to invoke this method
     *       whenever an {@link NSEApiException} is thrown anywhere in the controller or service layer.</li>
     * </ul>
     *
     * @param ex Caught {@link NSEApiException}.
     * @return RFC-7807 {@link ProblemDetail} with upstream status code or 502 Bad Gateway.
     */
    @ExceptionHandler(NSEApiException.class)
    public ProblemDetail handleNSEApiException(NSEApiException ex) {
        ProblemDetail problem = ProblemDetail.forStatusAndDetail(
                HttpStatus.valueOf(ex.getStatusCode() >= 400 && ex.getStatusCode() < 600 ? ex.getStatusCode() : 502),
                ex.getMessage()
        );
        problem.setTitle("NSE Upstream Integration Error");
        problem.setType(URI.create("https://errors.stockmarket.internal/nse-api-error"));
        problem.setProperty("timestamp", Instant.now());
        return problem;
    }

    /**
     * Handles Resilience4j rate limit exhaustion when client request volumes exceed configured thresholds.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@ExceptionHandler(RequestNotPermitted.class)</b>: Catches Resilience4j rate limiter rejections.</li>
     * </ul>
     *
     * @param ex Caught {@link RequestNotPermitted}.
     * @return RFC-7807 {@link ProblemDetail} with HTTP 429 Too Many Requests.
     */
    @ExceptionHandler(RequestNotPermitted.class)
    public ProblemDetail handleRateLimitException(RequestNotPermitted ex) {
        ProblemDetail problem = ProblemDetail.forStatusAndDetail(
                HttpStatus.TOO_MANY_REQUESTS,
                "NSE Market Gateway rate limit exceeded. Please throttle outgoing queries."
        );
        problem.setTitle("Rate Limit Exceeded");
        problem.setType(URI.create("https://errors.stockmarket.internal/rate-limit-exceeded"));
        problem.setProperty("timestamp", Instant.now());
        return problem;
    }

    /**
     * Handles Resilience4j circuit breaker trips when the circuit is in the OPEN state.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@ExceptionHandler(CallNotPermittedException.class)</b>: Catches circuit breaker short-circuits.</li>
     * </ul>
     *
     * @param ex Caught {@link CallNotPermittedException}.
     * @return RFC-7807 {@link ProblemDetail} with HTTP 503 Service Unavailable.
     */
    @ExceptionHandler(CallNotPermittedException.class)
    public ProblemDetail handleCircuitBreakerException(CallNotPermittedException ex) {
        ProblemDetail problem = ProblemDetail.forStatusAndDetail(
                HttpStatus.SERVICE_UNAVAILABLE,
                "NSE Gateway Circuit Breaker is OPEN due to upstream failure or bot blocking. Serving fallback/cached responses."
        );
        problem.setTitle("Circuit Breaker Active");
        problem.setType(URI.create("https://errors.stockmarket.internal/circuit-open"));
        problem.setProperty("timestamp", Instant.now());
        return problem;
    }

    /**
     * Handles Jakarta Bean Validation constraint violations (e.g. blank symbol or invalid request body).
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@ExceptionHandler(MethodArgumentNotValidException.class)</b>: Catches {@code @Valid} failures.</li>
     * </ul>
     *
     * @param ex Caught {@link MethodArgumentNotValidException}.
     * @return RFC-7807 {@link ProblemDetail} with HTTP 400 Bad Request and validation error field.
     */
    @ExceptionHandler(MethodArgumentNotValidException.class)
    public ProblemDetail handleValidationException(MethodArgumentNotValidException ex) {
        String errorMsg = ex.getBindingResult().getFieldError() != null
                ? ex.getBindingResult().getFieldError().getDefaultMessage()
                : "Validation failed";
        ProblemDetail problem = ProblemDetail.forStatusAndDetail(
                HttpStatus.BAD_REQUEST,
                "Validation failed for argument: " + errorMsg
        );
        problem.setTitle("Invalid Request Argument");
        problem.setType(URI.create("https://errors.stockmarket.internal/validation-error"));
        problem.setProperty("timestamp", Instant.now());
        return problem;
    }

    /**
     * Catch-all handler for unexpected internal server errors.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@ExceptionHandler(Exception.class)</b>: Catches any uncaught JVM or runtime exception.</li>
     * </ul>
     *
     * @param ex Unhandled {@link Exception}.
     * @return RFC-7807 {@link ProblemDetail} with HTTP 500 Internal Server Error.
     */
    @ExceptionHandler(Exception.class)
    public ProblemDetail handleGenericException(Exception ex) {
        ProblemDetail problem = ProblemDetail.forStatusAndDetail(
                HttpStatus.INTERNAL_SERVER_ERROR,
                ex.getMessage() != null ? ex.getMessage() : "An unexpected internal error occurred"
        );
        problem.setTitle("Internal Server Error");
        problem.setType(URI.create("https://errors.stockmarket.internal/internal-server-error"));
        problem.setProperty("timestamp", Instant.now());
        return problem;
    }
}
