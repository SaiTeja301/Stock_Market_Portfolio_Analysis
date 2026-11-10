# Architectural Reference Guide: Dual-Engine Generative AI Architecture (Spring AI & LangChain4j)

**Authoritative Architecture Document**  
**Service:** `nse-market-gateway-service` (Port: 8081)  
**Environment:** Java 17/19+, Spring Boot 3.2.5, LangChain4j 0.36.2, Spring AI 1.0.0-M1, Ollama (Local LLMs)  
**Target Audience:** Quantitative Strategists, AI Engineers, Algorithmic Trading Developers  

---

## 1. Executive Summary & Architectural Motivation

In institutional quantitative finance and stock market algorithmic trading, artificial intelligence cannot operate as a purely conversational "black box". Financial market applications have three uncompromising non-functional requirements:
1. **Mathematical Invariance:** Algorithmic sizing, linear regression channels, and risk calculations must execute deterministically in high-speed, compiled bytecode (Java engines), never computed by probabilistic LLM next-token predictors.
2. **Data Provenance & Hallucination Elimination:** Every quoted price, VWAP, bid-ask depth, and corporate filing must be verified in real time from authoritative exchange endpoints (NSE). Hallucinated prices or fake earnings dates lead to catastrophic capital loss.
3. **Multi-Model Specialization:** Different trading tasks demand different neural architectures. Fast equity scans require sub-second small parameter models (e.g. Gemma 2 2B), while complex trade reconciliation and risk memos require Chain-of-Thought reasoning engines (e.g. DeepSeek-R1 1.5B) or advanced mathematical models (Qwen 2.5 3B).

To satisfy these requirements, `nse-market-gateway-service` implements a **Dual-Engine Generative AI Architecture** co-locating **Spring AI** and **LangChain4j**.

```mermaid
graph TD
    User([Trader / Webhook / Algo Router]) --> GatewayController[SpringAiAdvisoryController]
    
    subgraph Spring_AI_Engine [Engine 1: Spring AI Framework]
        GatewayController -->|Intent / Model URL| SpringAiService[SpringAiAdvisoryService]
        SpringAiService --> DynamicRouter{Dynamic Model Router}
        DynamicRouter -->|Reasoning Intent| DeepSeek[DeepSeek-R1 1.5b]
        DynamicRouter -->|Fast Summary Intent| Gemma[Gemma 2 2b]
        DynamicRouter -->|Sentiment Intent| Llama[Llama 3.2 3b]
        DynamicRouter -->|Quantitative TRM Intent| Qwen[Qwen 2.5 3b]
        SpringAiService -.->|Function Callbacks| MarketDataAiTools[MarketDataAiTools Bean]
    end

    subgraph LangChain4j_Engine [Engine 2: LangChain4j Autonomous Agent]
        GatewayController -->|Dynamic Proxy| AiServiceProxy[@AiService StockAdvisoryAiAgent]
        AiServiceProxy --> ChatMemory[ChatMemoryProvider: Session Isolated Memory]
        AiServiceProxy --> AgentLoop{LLM Autonomous Tool Dispatch Loop}
        AgentLoop -->|Live Depth & Quote| Tool1[@Tool LangChain4jMarketDataTools]
        AgentLoop -->|Announcements & Embargo| Tool2[@Tool getCorporateAnnouncements]
        AgentLoop -->|TRM Mathematical Engine| Tool3[@Tool LangChain4jTRMTools]
        AgentLoop -->|Structured Extraction| StructOutput[TrmStructuredAdvisory Record]
        AgentLoop -->|Token Telemetry| ResultWrapper[Result TokenUsage Metrics]
    end

    subgraph Deterministic_Core [Deterministic Java 17 Service Layer]
        Tool1 & Tool2 & MarketDataAiTools --> NseService[NSEMarketGatewayService]
        Tool3 --> TrmService[TRMAlgorithmService: Classic, V1, V2, V3]
        NseService --> NseWeb[NSE Public Exchange Endpoints]
    end
```

---

## 2. LangChain4j: Core Annotations & Production Mechanics

Our project utilizes the official `dev.langchain4j:langchain4j:0.36.2` alongside `dev.langchain4j:langchain4j-ollama-spring-boot-starter:0.36.2`.

### 2.1 `@AiService`
- **Fully Qualified Name:** `dev.langchain4j.service.spring.AiService`
- **Class Applied:** [`StockAdvisoryAiAgent.java`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service/src/main/java/com/stockmarket/gateway/ai/StockAdvisoryAiAgent.java)
- **Internal Mechanics:** During Spring container initialization, LangChain4j scans for interfaces annotated with `@AiService`. It synthesizes a Java runtime dynamic proxy implementing all declared methods. This proxy transparently intercepts method invocations, binds the configured `OllamaChatModel`, wires all Spring beans containing `@Tool` annotations, and manages conversation history via `ChatMemoryProvider`.
- **Real-Time Trading Scenario:** Eliminates 1,000+ lines of imperative HTTP client code, JSON serialization/deserialization, and prompt string concatenation. The developer writes pure declarative Java interfaces; LangChain4j handles agentic tool loop orchestration.

### 2.2 `@SystemMessage`
- **Fully Qualified Name:** `dev.langchain4j.service.SystemMessage`
- **Internal Mechanics:** Injected as the high-priority `system` role message in the conversation payload sent to the LLM.
- **Real-Time Trading Scenario:** Establishes regulatory boundaries, behavioral guardrails, and compliance directives (e.g. SEBI/SEC disclaimers, forbidding margin or derivative trading, and enforcing data provenance tags: `[OBSERVED]`, `[CALCULATED]`, `[ASSUMED]`).

### 2.3 `@UserMessage` & `@V`
- **Fully Qualified Names:** `dev.langchain4j.service.UserMessage`, `dev.langchain4j.service.V`
- **Internal Mechanics:** `@UserMessage` defines the prompt template representing the inquiry using Mustache/Handlebars syntax (`{{variableName}}`). `@V("variableName")` binds Java method arguments directly into the template slots.
- **Real-Time Trading Scenario:** Mitigates **Prompt Injection Attacks**. Trader comments or external portfolio notes are placed strictly within isolated variable placeholders rather than being concatenated into unvalidated raw prompt strings.

```java
@SystemMessage("You are a Senior Quantitative Analyst at an institutional investment firm.")
@UserMessage("""
        Perform a complete TRM quantitative evaluation for equity symbol: {{symbol}}.
        Trader Position & Portfolio Context: {{userContext}}
        """)
String analyzeWithPromptTemplate(@V("symbol") String symbol, @V("userContext") String userContext);
```

### 2.4 `@MemoryId` & `ChatMemoryProvider`
- **Fully Qualified Names:** `dev.langchain4j.service.MemoryId`, `dev.langchain4j.memory.chat.ChatMemoryProvider`
- **Configuration:** [`LangChain4jConfig.java`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service/src/main/java/com/stockmarket/gateway/config/LangChain4jConfig.java)
- **Internal Mechanics:** In stateful dialogues, `@MemoryId` identifies the unique session key (e.g. trader ID, desk ID, or UUID). LangChain4j asks the `ChatMemoryProvider` bean for that session's `MessageWindowChatMemory(10)`. It appends previous user and assistant turns to the context window while automatically pruning messages beyond the 10-message ceiling to prevent context overflow.
- **Real-Time Trading Scenario:** Interactive trading desk assistant. A trader asks:
  1. Turn 1: *"Evaluate RELIANCE under TRM V2."* (Agent calculates levels).
  2. Turn 2: *"What if the stock drops 5% below that stop loss?"* (Agent recalls RELIANCE context without requiring the trader to re-state the ticker or prior calculations).

### 2.5 `@Tool` & `@P`
- **Fully Qualified Names:** `dev.langchain4j.agent.tool.Tool`, `dev.langchain4j.agent.tool.P`
- **Classes Applied:** [`LangChain4jMarketDataTools.java`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service/src/main/java/com/stockmarket/gateway/ai/LangChain4jMarketDataTools.java), [`LangChain4jTRMTools.java`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service/src/main/java/com/stockmarket/gateway/ai/LangChain4jTRMTools.java)
- **Internal Mechanics:** LangChain4j translates `@Tool` methods and `@P` parameter descriptions into OpenAPI-style tool schemas sent to the LLM. When the LLM emits a tool-call token sequence, LangChain4j pauses generation, invokes the annotated Java method with the deserialized parameters, feeds the execution result back to the model as a `tool` role message, and resumes completion until a final answer is reached.
- **Real-Time Trading Scenario:** Prevents financial hallucination. When asked for a stock's 52-week range or circuit limits, the LLM cannot fabricate numbers; it must trigger `getLiveStockQuote(symbol)` and observe verified exchange data.

### 2.6 Strongly-Typed Structured Record Extraction
- **Classes Applied:** [`TrmStructuredAdvisory.java`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service/src/main/java/com/stockmarket/gateway/dto/TrmStructuredAdvisory.java)
- **Annotations Used:** `dev.langchain4j.model.output.structured.Description`, `@JsonProperty`
- **Internal Mechanics:** When an `@AiService` method declares a Java Record return type, LangChain4j analyzes the Record schema via reflection, prompts the model to respond strictly in JSON conforming to that schema, and deserializes the JSON string directly into the Java Record without regex.
- **Real-Time Trading Scenario:** Algorithmic Order Management Systems (OMS). Downstream algorithmic routers cannot parse free-form prose. They require exact doubles and booleans: `trmStopLoss = 2860.00`, `trmTarget1 = 3020.00`, `binaryEventEmbargoClear = true`.

### 2.7 `Result<T>` & Token Telemetry
- **Fully Qualified Name:** `dev.langchain4j.service.Result`
- **Internal Mechanics:** Wraps the returned completion payload alongside execution telemetry: `tokenUsage().inputTokenCount()`, `tokenUsage().outputTokenCount()`, `tokenUsage().totalTokenCount()`, and `finishReason()`.
- **Real-Time Trading Scenario:** Enterprise GPU cluster capacity planning and billing telemetry. Ensures requests do not exceed local model context windows and tracks inference costs.

---

## 3. Spring AI: Architecture, Routing & Function Callbacks

Our project uses **Spring AI 1.0.0-M1** with Ollama auto-configuration.

### 3.1 `ChatClient` vs `ChatModel`
- **Configuration:** [`SpringAiConfig.java`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service/src/main/java/com/stockmarket/gateway/config/SpringAiConfig.java)
- **Mechanics:**
  - `ChatModel`: Low-level Spring AI SPI (Service Provider Interface) responsible for protocol-level interaction with Ollama.
  - `ChatClient`: Fluent, conversational client API built on top of `ChatModel`, supporting method chaining (`.prompt()`, `.system()`, `.user()`, `.call()`, `.content()`).
- **Real-Time Scenario:** In [`SpringAiAdvisoryService.java`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service/src/main/java/com/stockmarket/gateway/service/SpringAiAdvisoryService.java), `ChatClient` is used for high-level prompt assembly, while `ChatModel` is dynamically targeted during model overriding.

### 3.2 Dynamic Intent-Based Model Routing
Institutional trading desks require different models for different trading goals. `SpringAiAdvisoryService.resolveModel()` implements a 4-tier routing strategy:

| Target Model | Tier / Intent Key | Parameters | VRAM Usage | Latency | Real-Time Production Role |
|---|---|---|---|---|---|
| **`DeepSeek-R1:1.5b`** | `AUTO_ROUTED_FOR_DEEP_THINKING` | 1.5B | ~1.1 GB | ~3.5s | Chain-of-Thought mathematical breakdown, stop-loss derivations |
| **`qwen2.5:3b`** | `DEFAULT_QUANTITATIVE_TRM` | 3.0B | ~2.0 GB | ~1.8s | Multi-timeframe trend alignment, regression channel analysis |
| **`gemma2:2b`** | `AUTO_ROUTED_FOR_FAST_SUMMARY` | 2.6B | ~1.6 GB | ~0.8s | Rapid sub-second pre-trade screening and circuit audits |
| **`llama3.2:3b`** | `AUTO_ROUTED_FOR_SENTIMENT` | 3.2B | ~2.0 GB | ~1.6s | Market narrative, filing tone analysis, and news sentiment |

### 3.3 Spring AI Function Callbacks
- **Class Applied:** [`MarketDataAiTools.java`](file:///E:/Stock%20Market/Stock_Market_Portfolio_Analysis/Services/nse_market_gateway_service/src/main/java/com/stockmarket/gateway/tools/MarketDataAiTools.java)
- **Annotations Used:** `@Bean`, `@Description` on standard `java.util.function.Function<Request, Response>`.
- **Mechanics:** Spring AI automatically registers these beans as tool callbacks available to the `ChatClient` execution chain.

---

## 4. Framework Comparison: LangChain4j vs Spring AI

| Architectural Feature | LangChain4j (`dev.langchain4j`) | Spring AI (`org.springframework.ai`) | Used In This Project For |
|---|---|---|---|
| **Primary Interaction Paradigm** | Declarative Dynamic Interface Proxies (`@AiService`) | Fluent Builder API (`ChatClient`) & Model SPI (`ChatModel`) | Both (LangChain4j for autonomous agents; Spring AI for routing & service layer) |
| **Tool Calling Definition** | Method-level `@Tool` and parameter-level `@P` | `@Bean Function<Req, Resp>` with `@Description` | Both (LangChain4j tool beans; Spring AI function beans) |
| **Session Chat Memory** | Built-in `@MemoryId` + `ChatMemoryProvider` | Custom advisor / external memory store | LangChain4j (`MessageWindowChatMemory`) |
| **Structured Output Extraction** | Direct Java Record / POJO return types with JSON schema | `BeanOutputParser<T>` | LangChain4j (`TrmStructuredAdvisory`) |
| **Token Telemetry** | `Result<T>` with `TokenUsage` metrics | `ChatResponse.getMetadata().getUsage()` | Both |
| **Multi-Model Dynamic Routing** | Requires multiple custom `@AiService` instances | Native model options override per request | Spring AI (`SpringAiAdvisoryService`) |

---

## 5. Complete 30-Endpoint REST API Reference

All endpoints run on base URL: `http://localhost:8081`

### Section 1: Real-Time Market Data Endpoints
1. `GET  /api/v1/market/quote/{symbol}` - Live quote, LTP, VWAP, 52W range, circuit limits & 5-level DOM orderbook.
2. `GET  /api/v1/market/charts/{symbol}?days=1D` - Intraday and historical time-series price data.
3. `GET  /api/v1/market/announcements/{symbol}` - Recent corporate announcements, filings, and board meetings.
4. `GET  /api/v1/market/top-ten` - Top 10 advances, declines, and most active NSE equities.
5. `GET  /api/v1/market/yearwise/{symbol}` - Multi-period returns (1D, 1W, 1M, 1Y, 3Y, 5Y) vs Nifty 50.
6. `GET  /api/v1/market/indices/{symbol}` - Sectoral and benchmark indices containing the stock.
7. `GET  /api/v1/market/search?query={query}` - Fast fuzzy equity search by ticker, company name, or ISIN.

### Section 2: Generative AI Advisory & Autonomous Agent Endpoints
8. `GET  /api/v1/ai/models` - Catalog of validated local Ollama LLMs with specialization tiers and VRAM ratings.
9. `POST /api/v1/ai/analyze-stock` - Spring AI advisory analysis with automated intent-based model routing.
10. `POST /api/v1/ai/analyze-stock/{model}` - Spring AI analysis pinned to an explicit URL model parameter.
11. `GET  /api/v1/ai/catalyst-check/{symbol}` - 48-hour binary event risk embargo check via Spring AI.
12. `POST /api/v1/ai/langchain4j/analyze-stock` - LangChain4j autonomous agent with declarative `@Tool` callbacks.
13. `POST /api/v1/ai/langchain4j/analyze-templated` - Parameterized template analysis via `@UserMessage` & `@V`.
14. `POST /api/v1/ai/langchain4j/chat/{sessionId}` - Stateful multi-turn conversational chat with `@MemoryId`.
15. `POST /api/v1/ai/langchain4j/analyze-structured` - Strongly-typed structured JSON extraction directly to Java Record (`TrmStructuredAdvisory`).
16. `GET  /api/v1/ai/langchain4j/token-metadata/{symbol}` - Telemetry & token usage breakdown via `Result<T>`.
17. `GET  /api/v1/ai/langchain4j/catalyst-check/{symbol}` - Binary event risk audit via LangChain4j autonomous agent.

### Section 3: Teja Recovery Model (TRM) Algorithmic Engine Endpoints
18. `POST /api/v1/trm/classic` - TRM Classic: Recursive cumulative capital allocation ($B_{n+1} = 2 \cdot T_n$).
19. `GET  /api/v1/trm/classic/{symbol}` - TRM Classic calculated from live stock quote.
20. `POST /api/v1/trm/v1` - TRM V1: Percentage-triggered stock equity model ($P_n = P_0 \cdot (1 - L)^n$).
21. `GET  /api/v1/trm/v1/{symbol}` - TRM V1 calculated from live stock quote.
22. `POST /api/v1/trm/v2` - TRM V2: Actual-portfolio-loss adaptive model (99.68% capital savings vs V1).
23. `GET  /api/v1/trm/v2/{symbol}` - TRM V2 calculated from live stock quote.
24. `POST /api/v1/trm/v3` - TRM V3: Asymmetric Dual-Path (Profit Runner Vaulting + Hard-Stop Recovery Surcharge).
25. `GET  /api/v1/trm/v3/{symbol}` - TRM V3 calculated from live stock quote.
26. `POST /api/v1/trm/evaluate-all` - Concurrent multi-model evaluation across all 4 TRM algorithms.
27. `GET  /api/v1/trm/evaluate/{symbol}` - Live 360° Confluence Evaluation: All 4 TRMs + 100-bar OLS Regression Channel.

### Section 4: System Health, Metrics & Actuator Endpoints
28. `GET  /actuator/health` - Health status and subsystem indicators (JVM, disk space, cache).
29. `GET  /actuator/info` - Application version and build metadata.
30. `GET  /actuator/metrics` - Published JVM, GC, HTTP request latency, and Caffeine cache metrics.

---

## 6. Build & Test Verification

To execute the entire test suite on Windows without native memory fragmentation:

```powershell
cmd /c "set MAVEN_OPTS=-Xmx512m && mvn test"
```

### Verification Results:
- **Total Tests:** 35
- **Failures:** 0
- **Errors:** 0
- **Skipped:** 0
- **Test Execution Classes:**
  1. `SpringAiAdvisoryControllerTest` (5 tests covering chat memory, structured records, token metrics, models, and analysis)
  2. `SpringAiAdvisoryServiceTest` (4 tests covering intent routing and TRM deterministic evaluation)
  3. `LangChain4jMarketDataToolsTest` (4 tests covering live quotes, announcements, and circuit limits)
  4. `TRMAlgorithmControllerTest` (6 tests covering TRM REST endpoints)
  5. `TRMAlgorithmServiceTest` (8 tests covering TRM algorithms and invariants)
  6. `NSEMarketGatewayControllerTest` (1 test covering market data endpoints)
  7. `NSEMarketGatewayServiceTest` (7 tests covering data fetching, fallback, and resilience)
