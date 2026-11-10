package com.stockmarket.gateway.controller;

import com.stockmarket.gateway.dto.AiStockAnalysisRequest;
import com.stockmarket.gateway.dto.AiStockAnalysisResponse;
import com.stockmarket.gateway.service.SpringAiAdvisoryService;
import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import com.stockmarket.gateway.ai.StockAdvisoryAiAgent;
import com.stockmarket.gateway.dto.ModelMetadataRecord;
import com.stockmarket.gateway.dto.TrmStructuredAdvisory;
import dev.langchain4j.service.Result;
import org.springframework.lang.Nullable;

import java.time.Instant;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * ====================================================================================================
 *         DUAL-ENGINE AI CONTROLLER: SPRING AI & LANGCHAIN4J HYBRID ADVISORY SUITE
 * ====================================================================================================
 *
 * Architectural Purpose:
 * Exposes an enterprise-grade REST API bridging two state-of-the-art Generative AI frameworks:
 * 1. Spring AI:
 *    - Fluent {@code ChatClient} API and lower-level {@code ChatModel} abstraction.
 *    - Dynamic multi-model routing (Qwen 2.5 3B, DeepSeek-R1 1.5B, Gemma 2 2B, Llama 3.2 3B).
 *    - Spring-native Function Callbacks ({@code @Bean Function<Req, Resp>} with {@code @Description}).
 * 2. LangChain4j:
 *    - Declarative {@code @AiService} dynamic interface proxies.
 *    - Automated tool-dispatch loop ({@code @Tool} and {@code @P} annotations).
 *    - Stateful multi-turn conversational chat memory ({@code @MemoryId} + {@code ChatMemoryProvider}).
 *    - Strongly-typed structured JSON extraction directly into Java Records ({@link TrmStructuredAdvisory}).
 *    - Token telemetry and observability ({@code Result<T>} capturing token counts and finish reasons).
 *
 * Annotations Applied on this Class:
 * <ul>
 *   <li><b>@RestController</b>: Specialized Spring component annotation combining {@code @Controller}
 *       and {@code @ResponseBody}. It tells Spring MVC that this class handles incoming HTTP requests
 *       and automatically serializes Java return values directly into HTTP response bodies as JSON
 *       using Jackson ObjectMapper.</li>
 *   <li><b>@RequestMapping("/api/v1/ai")</b>: Defines the base URI path prefix for all endpoints
 *       exposed by this controller. Real-time scenario: Versioned API segregation (/api/v1/ai) ensuring
 *       backward compatibility across algorithmic trading microservices.</li>
 *   <li><b>@RequiredArgsConstructor</b>: Lombok annotation generating a constructor for all {@code final}
 *       fields ({@code aiAdvisoryService} and {@code stockAdvisoryAiAgent}). Eliminates field injection
 *       in favor of robust, testable constructor injection recommended by Spring team.</li>
 *   <li><b>@Slf4j</b>: Lombok annotation creating a production-grade SLF4J Logger instance named {@code log},
 *       enabling low-latency asynchronous structured logging for trade execution audits.</li>
 * </ul>
 */
@Slf4j
@RestController
@RequestMapping("/api/v1/ai")
@RequiredArgsConstructor
public class SpringAiAdvisoryController {

    private final SpringAiAdvisoryService aiAdvisoryService;

    /**
     * Optional injection of the LangChain4j autonomous agent.
     * Annotated with {@code @Nullable} to ensure the microservice starts seamlessly even in
     * environments where LangChain4j Ollama bindings are offline or bypassed during unit tests.
     */
    @Nullable
    private final StockAdvisoryAiAgent stockAdvisoryAiAgent;

    // ================================================================================================
    //                                    SPRING AI REST ENDPOINTS
    // ================================================================================================

    /**
     * Retrieves the curated catalog of validated local LLM models available on the Ollama instance.
     *
     * Annotation Explained:
     * <ul>
     *   <li><b>@GetMapping("/models")</b>: Maps HTTP GET requests to "/api/v1/ai/models".
     *       Used for idempotent retrieval of the active model portfolio without modifying state.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Front-end trading dashboards and algorithmic execution engines call this endpoint upon startup
     * to populate model selection dropdowns (e.g. choosing DeepSeek-R1 for complex macro synthesis
     * vs Gemma 2 for sub-second rapid equity filtering).
     *
     * @return HTTP 200 OK with the list of {@link ModelMetadataRecord} entries.
     */
    @GetMapping("/models")
    public ResponseEntity<List<ModelMetadataRecord>> getModels() {
        log.info("Fetching available local Ollama text and reasoning models");
        return ResponseEntity.ok(aiAdvisoryService.getAvailableModels());
    }

    /**
     * Evaluates an NSE equity symbol using Spring AI's dynamic intent router and deterministic TRM engine.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@PostMapping("/analyze-stock")</b>: Maps HTTP POST requests to "/api/v1/ai/analyze-stock".
     *       POST is chosen because the analysis payload contains structured constraints and user context.</li>
     *   <li><b>@Valid</b>: Triggers Jakarta Bean Validation on the incoming {@link AiStockAnalysisRequest}
     *       (verifying non-blank ticker symbol and valid format). If validation fails, Spring throws
     *       {@code MethodArgumentNotValidException} and returns HTTP 400 Bad Request before hitting business logic.</li>
     *   <li><b>@RequestBody</b>: Directs Spring's {@code HttpMessageConverter} to deserialize the incoming
     *       HTTP JSON request body into the strongly-typed Java Record {@link AiStockAnalysisRequest}.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Institutional portfolio rebalancing check. A trader submits an equity symbol with portfolio notes,
     * and the service autonomously inspects real-time NSE quotes, checks corporate filings for upcoming
     * binary events, calculates TRM multi-timeframe indicators, and generates an executive recommendation.
     *
     * @param request Validated equity analysis request record.
     * @return HTTP 200 OK with {@link AiStockAnalysisResponse} containing TRM scores, targets, and Markdown report.
     */
    @PostMapping("/analyze-stock")
    public ResponseEntity<AiStockAnalysisResponse> analyzeStock(@Valid @RequestBody AiStockAnalysisRequest request) {
        log.info("Received Spring AI stock analysis request for [{}]", request.symbol());
        return ResponseEntity.ok(aiAdvisoryService.analyzeStock(request));
    }

    /**
     * Evaluates an NSE equity symbol while forcing execution through an explicitly requested LLM model.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@PostMapping("/analyze-stock/{model}")</b>: Parameterized URI template allowing URL-driven
     *       model selection (e.g. "/api/v1/ai/analyze-stock/deepseek" or "/api/v1/ai/analyze-stock/gemma").</li>
     *   <li><b>@PathVariable String model</b>: Extracts the dynamic URI segment {@code {model}} directly
     *       into the method parameter.</li>
     *   <li><b>@Valid @RequestBody</b>: Validates and deserializes the incoming analysis parameters.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * A quantitative researcher wants to perform A/B model benchmarking on the same equity chart,
     * comparing the Chain-of-Thought reasoning of DeepSeek-R1 against the rapid summarization of Gemma 2.
     *
     * @param model Target model family or alias (e.g. "deepseek", "qwen", "gemma", "llama").
     * @param request Validated equity analysis request record.
     * @return HTTP 200 OK with model-specific {@link AiStockAnalysisResponse}.
     */
    @PostMapping("/analyze-stock/{model}")
    public ResponseEntity<AiStockAnalysisResponse> analyzeStockWithModel(
            @PathVariable String model,
            @Valid @RequestBody AiStockAnalysisRequest request) {
        AiStockAnalysisRequest overriddenRequest = new AiStockAnalysisRequest(
                request.symbol(),
                request.analysisType(),
                model,
                request.userContext()
        );
        log.info("Received Spring AI stock analysis request for [{}] with URL model [{}]", request.symbol(), model);
        return ResponseEntity.ok(aiAdvisoryService.analyzeStock(overriddenRequest));
    }

    /**
     * Audits the 48-Hour Binary Event Risk Embargo for a given NSE ticker symbol via Spring AI.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/catalyst-check/{symbol}")</b>: Maps HTTP GET requests to inspect corporate
     *       events and regulatory announcements for a specific ticker symbol.</li>
     *   <li><b>@PathVariable String symbol</b>: Injects the URL path parameter (e.g. "RELIANCE") into the method.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * TRM Algorithm Pre-Flight Safety Check. Before executing any swing or momentum entry, the trading system
     * MUST verify that the stock does not have an imminent earnings release, board meeting, or regulatory
     * inquiry within the next 48 hours, which could introduce unpredictable binary volatility.
     *
     * @param symbol NSE equity ticker symbol (e.g. "INFY", "TCS", "TATAMOTORS").
     * @return HTTP 200 OK with a catalyst clearance report and embargo status.
     */
    @GetMapping("/catalyst-check/{symbol}")
    public ResponseEntity<Map<String, String>> checkCatalysts(@PathVariable String symbol) {
        log.info("Checking binary event catalyst risk for [{}] via Spring AI", symbol);
        String report = aiAdvisoryService.checkBinaryEventRisk(symbol);
        return ResponseEntity.ok(Map.of(
                "symbol", symbol.toUpperCase(),
                "catalystRiskReport", report
        ));
    }

    // ================================================================================================
    //                                 LANGCHAIN4J REST ENDPOINTS
    // ================================================================================================

    /**
     * Executes autonomous stock analysis via LangChain4j's declarative {@code @AiService} with dynamic {@code @Tool} callbacks.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@PostMapping("/langchain4j/analyze-stock")</b>: Dedicated endpoint routing requests to the
     *       LangChain4j dynamic agent proxy ({@link StockAdvisoryAiAgent}).</li>
     *   <li><b>@Valid @RequestBody</b>: Validates and maps JSON input payload.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Autonomous multi-step investigation. The LLM receives the user's trading goal, realizes it lacks
     * current price and volume metrics, autonomously invokes the {@code getLiveStockQuote} tool, observes
     * that the stock is near its 100-period channel floor, invokes {@code evaluateTRMModels} to calculate
     * regression slope, checks corporate filings, and synthesizes a final investment memo.
     *
     * @param request Analysis request containing ticker symbol and optional trader context.
     * @return HTTP 200 OK with synthesized analysis, provenance tags ([OBSERVED], [CALCULATED], [ASSUMED]), and timestamp.
     */
    @PostMapping("/langchain4j/analyze-stock")
    public ResponseEntity<Map<String, Object>> analyzeWithLangChain4j(@Valid @RequestBody AiStockAnalysisRequest request) {
        log.info("Received LangChain4j @AiService stock analysis request for [{}]", request.symbol());
        if (stockAdvisoryAiAgent == null) {
            AiStockAnalysisResponse fallback = aiAdvisoryService.analyzeStock(request);
            return ResponseEntity.ok(Map.of(
                    "symbol", request.symbol().toUpperCase(),
                    "agentFramework", "Spring AI Fallback",
                    "analysis", fallback.fullAnalysisMarkdown(),
                    "recommendation", fallback.recommendation(),
                    "setupScore", fallback.trmSetupScore(),
                    "timestamp", Instant.now().toString()
            ));
        }
        String userPrompt = String.format("Evaluate symbol: %s under TRM Strategy confluence. User Context: %s",
                request.symbol(), request.userContext() != null ? request.userContext() : "Analyze entry setup");
        String result = stockAdvisoryAiAgent.analyzeStockWithTools(userPrompt);
        return ResponseEntity.ok(Map.of(
                "symbol", request.symbol().toUpperCase(),
                "agentFramework", "LangChain4j (@AiService + @Tool)",
                "analysis", result,
                "timestamp", Instant.now().toString()
        ));
    }

    /**
     * Executes parameterized stock analysis using LangChain4j's Handlebars-style template interpolation ({@code @V}).
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@PostMapping("/langchain4j/analyze-templated")</b>: Maps HTTP POST requests to trigger
     *       the {@code analyzeWithPromptTemplate} method on {@link StockAdvisoryAiAgent}.</li>
     *   <li><b>@Valid @RequestBody</b>: Ensures structured input validation.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Injection-safe standardized screening. By binding {@code request.symbol()} to {@code @V("symbol")}
     * and {@code request.userContext()} to {@code @V("userContext")}, the system prevents prompt injection
     * attacks from untrusted external user comments while maintaining strict prompt formatting.
     *
     * @param request Validated request payload.
     * @return HTTP 200 OK with templated analysis result.
     */
    @PostMapping("/langchain4j/analyze-templated")
    public ResponseEntity<Map<String, Object>> analyzeWithLangChain4jTemplate(@Valid @RequestBody AiStockAnalysisRequest request) {
        log.info("Received LangChain4j templated analysis request for [{}]", request.symbol());
        if (stockAdvisoryAiAgent == null) {
            AiStockAnalysisResponse fallback = aiAdvisoryService.analyzeStock(request);
            return ResponseEntity.ok(Map.of(
                    "symbol", request.symbol().toUpperCase(),
                    "agentFramework", "Spring AI Fallback",
                    "analysis", fallback.fullAnalysisMarkdown(),
                    "timestamp", Instant.now().toString()
            ));
        }
        String userContext = request.userContext() != null ? request.userContext() : "Standard TRM Entry Evaluation";
        String result = stockAdvisoryAiAgent.analyzeWithPromptTemplate(request.symbol().toUpperCase(), userContext);
        return ResponseEntity.ok(Map.of(
                "symbol", request.symbol().toUpperCase(),
                "agentFramework", "LangChain4j (@AiService + @UserMessage template + @V)",
                "analysis", result,
                "timestamp", Instant.now().toString()
        ));
    }

    /**
     * Conducts stateful multi-turn conversational dialogue with the AI agent using session-isolated memory ({@code @MemoryId}).
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@PostMapping("/langchain4j/chat/{sessionId}")</b>: Maps HTTP POST requests to manage
     *       conversational continuity keyed by a unique session identifier.</li>
     *   <li><b>@PathVariable String sessionId</b>: Injected into {@code @MemoryId String sessionId}.
     *       Under the hood, LangChain4j's {@code ChatMemoryProvider} retrieves the isolated rolling
     *       chat window (last 10 messages) specifically for this trader session.</li>
     *   <li><b>@RequestBody Map<String, String> payload</b>: Extracts the user's conversational message
     *       from JSON payload (e.g. {@code {"message": "What is the stop loss we calculated?"}}).</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * An interactive trading assistant UI. The user first asks "Analyze HDFCBANK". The agent responds with
     * calculations. Two minutes later, the user asks "What if it breaks below 1700?". Because of
     * {@code @MemoryId}, the agent remembers the previous HDFCBANK discussion without requiring the user
     * to re-specify the stock symbol or prior price context.
     *
     * @param sessionId Unique session ID (e.g. user ID, desk ID, or UUID).
     * @param payload Map containing the "message" key.
     * @return HTTP 200 OK with conversational response and session metadata.
     */
    @PostMapping("/langchain4j/chat/{sessionId}")
    public ResponseEntity<Map<String, Object>> chatWithLangChain4j(
            @PathVariable String sessionId,
            @RequestBody Map<String, String> payload) {
        String userMessage = payload.getOrDefault("message", payload.getOrDefault("prompt", "Hello"));
        log.info("Received LangChain4j conversational chat message for session [{}]", sessionId);

        if (stockAdvisoryAiAgent == null) {
            return ResponseEntity.ok(Map.of(
                    "sessionId", sessionId,
                    "agentFramework", "Spring AI Fallback (Memory Offline)",
                    "response", "LangChain4j agent bean is offline. Please verify Ollama connectivity.",
                    "timestamp", Instant.now().toString()
            ));
        }

        String response = stockAdvisoryAiAgent.conversationalAdvisory(sessionId, userMessage);
        return ResponseEntity.ok(Map.of(
                "sessionId", sessionId,
                "agentFramework", "LangChain4j (@MemoryId + ChatMemoryProvider)",
                "response", response,
                "timestamp", Instant.now().toString()
        ));
    }

    /**
     * Extracts strongly-typed structured JSON directly into a Java Record ({@link TrmStructuredAdvisory}) without regex.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@PostMapping("/langchain4j/analyze-structured")</b>: Maps HTTP POST requests for structured
     *       type-safe algorithmic analysis.</li>
     *   <li><b>@Valid @RequestBody</b>: Validates the incoming equity request.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Automated Algorithmic Order Management System (OMS). Downstream algorithmic trading systems cannot
     * parse prose or markdown. They require clean, strongly-typed JSON with double/float values
     * (e.g. {@code trmStopLoss = 2860.0, trmTarget1 = 3020.0, setupScore = 85, recommendation = "STRONG_BUY"}).
     * LangChain4j dynamically enforces JSON schema adherence and deserializes the output directly
     * into the Java {@link TrmStructuredAdvisory} Record.
     *
     * @param request Equity request record.
     * @return HTTP 200 OK with typed {@link TrmStructuredAdvisory} record.
     */
    @PostMapping("/langchain4j/analyze-structured")
    public ResponseEntity<TrmStructuredAdvisory> analyzeStockStructured(@Valid @RequestBody AiStockAnalysisRequest request) {
        log.info("Received LangChain4j structured output extraction request for [{}]", request.symbol());
        if (stockAdvisoryAiAgent == null) {
            // Provide a graceful fallback Record when the agent bean is inactive
            AiStockAnalysisResponse fallback = aiAdvisoryService.analyzeStock(request);
            double estimatedLtp = fallback.invalidationStopLoss() > 0 ? Math.round(fallback.invalidationStopLoss() / 0.96 * 100.0) / 100.0 : 0.0;
            boolean embargoClear = fallback.binaryEventRisk() == null || !fallback.binaryEventRisk().toUpperCase().contains("EMBARGO");
            TrmStructuredAdvisory fallbackRecord = new TrmStructuredAdvisory(
                    request.symbol().toUpperCase(),
                    estimatedLtp,
                    fallback.invalidationStopLoss(),
                    fallback.target1(),
                    fallback.target2(),
                    fallback.trmSetupScore(),
                    fallback.recommendation(),
                    fallback.executiveSummary(),
                    embargoClear
            );
            return ResponseEntity.ok(fallbackRecord);
        }

        String userContext = request.userContext() != null ? request.userContext() : "Extract exact numeric levels";
        TrmStructuredAdvisory advisory = stockAdvisoryAiAgent.analyzeStockStructured(request.symbol().toUpperCase(), userContext);
        return ResponseEntity.ok(advisory);
    }

    /**
     * Executes analysis and returns execution metadata alongside token consumption metrics using {@link Result}.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/langchain4j/token-metadata/{symbol}")</b>: Maps HTTP GET requests for token telemetry.</li>
     *   <li><b>@PathVariable String symbol</b>: The ticker symbol to evaluate.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Institutional cost auditing and inference observability. In production, tracking prompt token counts,
     * completion token counts, latency, and LLM finish reasons (e.g. "STOP" vs "LENGTH") is mandatory
     * to manage GPU server load and ensure prompts do not overflow model context windows.
     *
     * @param symbol NSE equity symbol.
     * @return HTTP 200 OK with content, token usage breakdown, and finish reason.
     */
    @GetMapping("/langchain4j/token-metadata/{symbol}")
    public ResponseEntity<Map<String, Object>> analyzeWithResultMetadata(@PathVariable String symbol) {
        log.info("Received LangChain4j token telemetry request for [{}]", symbol);
        if (stockAdvisoryAiAgent == null) {
            return ResponseEntity.ok(Map.of(
                    "symbol", symbol.toUpperCase(),
                    "agentFramework", "Spring AI Fallback",
                    "content", "LangChain4j agent offline. Token telemetry unavailable.",
                    "totalTokens", 0,
                    "timestamp", Instant.now().toString()
            ));
        }

        Result<String> result = stockAdvisoryAiAgent.analyzeWithResultMetadata(symbol.toUpperCase());
        Map<String, Object> response = new HashMap<>();
        response.put("symbol", symbol.toUpperCase());
        response.put("agentFramework", "LangChain4j (Result<T> Telemetry)");
        response.put("content", result.content());

        if (result.tokenUsage() != null) {
            response.put("inputTokenCount", result.tokenUsage().inputTokenCount());
            response.put("outputTokenCount", result.tokenUsage().outputTokenCount());
            response.put("totalTokenCount", result.tokenUsage().totalTokenCount());
        } else {
            response.put("totalTokenCount", "N/A (Local Ollama)");
        }

        if (result.finishReason() != null) {
            response.put("finishReason", result.finishReason().name());
        }

        response.put("timestamp", Instant.now().toString());
        return ResponseEntity.ok(response);
    }

    /**
     * Performs specialized regulatory and catalyst surveillance via LangChain4j's autonomous tool loop.
     *
     * Annotations Explained:
     * <ul>
     *   <li><b>@GetMapping("/langchain4j/catalyst-check/{symbol}")</b>: Maps HTTP GET requests for catalyst auditing.</li>
     *   <li><b>@PathVariable String symbol</b>: The equity ticker symbol to audit.</li>
     * </ul>
     *
     * Real-Time Trading Scenario:
     * Automated pre-trade risk compliance. LangChain4j invokes {@code getCorporateAnnouncements}
     * tool to inspect upcoming board meetings, dividend notices, and corporate actions, outputting
     * a regulatory clearance memo to prevent trading ahead of binary events.
     *
     * @param symbol The NSE equity ticker symbol.
     * @return HTTP 200 OK with catalyst risk clearance report.
     */
    @GetMapping("/langchain4j/catalyst-check/{symbol}")
    public ResponseEntity<Map<String, String>> checkCatalystWithLangChain4j(@PathVariable String symbol) {
        log.info("Checking binary event risk via LangChain4j @AiService for [{}]", symbol);
        if (stockAdvisoryAiAgent == null) {
            return ResponseEntity.ok(Map.of(
                    "symbol", symbol.toUpperCase(),
                    "agentFramework", "Spring AI Fallback",
                    "catalystRiskReport", aiAdvisoryService.checkBinaryEventRisk(symbol)
            ));
        }
        String report = stockAdvisoryAiAgent.auditCatalystRiskWithTools("Audit binary event risk for symbol: " + symbol.toUpperCase());
        return ResponseEntity.ok(Map.of(
                "symbol", symbol.toUpperCase(),
                "agentFramework", "LangChain4j (@AiService + @Tool)",
                "catalystRiskReport", report
        ));
    }
}
