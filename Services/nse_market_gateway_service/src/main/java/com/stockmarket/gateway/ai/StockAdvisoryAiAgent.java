package com.stockmarket.gateway.ai;

import com.stockmarket.gateway.dto.TrmStructuredAdvisory;
import dev.langchain4j.service.MemoryId;
import dev.langchain4j.service.Result;
import dev.langchain4j.service.SystemMessage;
import dev.langchain4j.service.UserMessage;
import dev.langchain4j.service.V;
import dev.langchain4j.service.spring.AiService;

/**
 * ====================================================================================================
 *                 LANGCHAIN4J DECLARATIVE AI SERVICE: STOCK ADVISORY AGENT
 * ====================================================================================================
 * 
 * Architectural Overview:
 * This interface is the centerpiece of the LangChain4j integration in this microservice.
 * Rather than writing boilerplate HTTP clients, manual JSON deserializers, and prompt string builders,
 * LangChain4j uses Spring dynamic proxies to turn this interface into a fully autonomous AI agent.
 * 
 * Key LangChain4j Annotations Explained in Depth:
 * 
 * 1. @AiService (dev.langchain4j.service.spring.AiService):
 *    - Purpose: Marks this interface as an AI Service component in the Spring context.
 *    - Internal Mechanics: During application context startup, LangChain4j's Spring Boot starter
 *      detects this annotation and generates a dynamic runtime proxy implementing all declared methods.
 *      It automatically wires:
 *        a) The configured ChatLanguageModel (OllamaChatModel).
 *        b) All Spring beans containing methods annotated with @Tool (e.g. LangChain4jMarketDataTools, LangChain4jTRMTools).
 *        c) The configured ChatMemoryProvider for stateful multi-turn dialogues.
 *    - Real-Time Scenario: Eliminates 1,000+ lines of imperative glue code. Whenever a controller or service
 *      calls a method on this interface, the request is transparently converted to an LLM dialogue, tools
 *      are called if needed, and responses are formatted according to the return type.
 * 
 * 2. @SystemMessage (dev.langchain4j.service.SystemMessage):
 *    - Purpose: Defines the foundational persona, behavioral guardrails, and operational boundaries for the LLM.
 *    - Internal Mechanics: Injected as the high-priority `system` role message in the chat conversation history.
 *    - Real-Time Scenario: Enforces regulatory compliance (SEBI/SEC risk disclaimers, forbidding margin trading,
 *      mandating strict data provenance tags [OBSERVED], [CALCULATED], [ASSUMED]).
 * 
 * 3. @UserMessage (dev.langchain4j.service.UserMessage):
 *    - Purpose: Defines the prompt template representing the user's inquiry.
 *    - Internal Mechanics: Supports Handlebars-style variable interpolation using Mustache syntax: {{variableName}}.
 *    - Real-Time Scenario: Prevents prompt injection by isolating user-supplied strings into controlled template slots.
 * 
 * 4. @V (dev.langchain4j.service.V):
 *    - Purpose: Binds a method parameter to a named variable placeholder inside a @UserMessage or @SystemMessage template.
 *    - Internal Mechanics: LangChain4j matches `@V("symbol") String ticker` to `{{symbol}}` in the template string.
 *    - Real-Time Scenario: Clean, type-safe API invocation where callers supply variables without string concatenation.
 * 
 * 5. @MemoryId (dev.langchain4j.service.MemoryId):
 *    - Purpose: Identifies the unique session or user conversation context for stateful multi-turn dialogues.
 *    - Internal Mechanics: LangChain4j queries the registered ChatMemoryProvider with this ID to load past
 *      messages and appends the latest turn before sending the combined window to the LLM.
 *    - Real-Time Scenario: Enables interactive conversational trading assistants where a trader asks
 *      "What is Reliance's stop loss?", followed by "What if it drops 5% more?", and the LLM retains context
 *      without confusing it with other active trader sessions.
 * 
 * 6. Structured Output Return Types (POJO / Record extraction):
 *    - Purpose: Direct extraction of typed Java objects (e.g. {@link TrmStructuredAdvisory}) without regex.
 *    - Internal Mechanics: LangChain4j constructs a strict JSON Schema from the return Record, forces the LLM
 *      into JSON output mode, and deserializes the JSON string into the typed instance.
 *    - Real-Time Scenario: Automated trade execution pipelines that require exact numerical values (floats, doubles, enums)
 *      to route orders directly to broker APIs.
 * 
 * 7. Result<T> (dev.langchain4j.service.Result):
 *    - Purpose: Wraps the response payload alongside execution metadata (token counts, finish reason, tool executions).
 *    - Real-Time Scenario: Enterprise auditing, API cost tracking, and token consumption analytics.
 */
@AiService
public interface StockAdvisoryAiAgent {

    /**
     * Executes autonomous stock analysis with full dynamic tool-calling.
     * 
     * Annotation: @SystemMessage
     * Instructs the LLM to act as a quantitative strategist, mandating live data tool calls before answering.
     * 
     * Annotation: @UserMessage
     * Designates the single string parameter as the complete user prompt.
     * 
     * @param userPrompt The free-form user question or analysis objective.
     * @return Formatted analytical Markdown containing data provenance tags and executive recommendation.
     */
    @SystemMessage("""
            You are an enterprise-grade Quantitative Stock Market Advisory AI Agent specializing in Indian Equities (NSE)
            and the Teja Recovery Model (TRM) Strategy.
            
            You have direct access to automated tools to inspect the real-time market:
            - getLiveStockQuote: Returns LTP, VWAP, Change %, 52W Range, and 5-level DOM Order Depth.
            - getCorporateAnnouncements: Returns latest corporate announcements, board meetings, and regulatory filings.
            - getHistoricalPerformance: Returns 1-day, 1-week, and 1-year returns against benchmark indices.
            - checkCircuitStatus: Checks whether a stock is locked in upper or lower circuit bands.
            - evaluateTRMModels: Evaluates all 4 TRM algorithmic editions (Classic, V1, V2, V3) and 100-period regression channel.
            
            MANDATORY INSTRUCTIONS:
            1. ALWAYS call tools to observe actual live data for the stock before answering. NEVER hallucinate prices, volume, or filings.
            2. Calculate TRM Stop Loss (typically 4% below entry/LTP), Target 1 (+5%), Target 2 (+10%).
            3. Verify that the 48-Hour Binary Event Embargo is CLEAR (no immediate board meetings, earnings releases, or delisting notices).
            4. Tag every statement with data provenance:
               - [OBSERVED]: For factual metrics directly returned by tools.
               - [CALCULATED]: For mathematically derived levels, ratios, and price targets.
               - [ASSUMED]: For risk parameters, trader assumptions, and scenario guidance.
            5. Conclude with a clear Executive Recommendation: STRONG BUY, TRM RE-ENTRY, HOLD / WAIT, PARTIAL EXIT, or AVOID.
            """)
    String analyzeStockWithTools(@UserMessage String userPrompt);

    /**
     * Executes parameterized stock analysis using Handlebars-style template interpolation with @V.
     * 
     * Annotations: @UserMessage with {{symbol}} and {{userContext}}, paired with @V("symbol") and @V("userContext").
     * Role: Clean template separation that prevents prompt injection while ensuring strict input structuring.
     * 
     * @param symbol The NSE equity ticker symbol (e.g. RELIANCE, TCS).
     * @param userContext Contextual details provided by the trader (e.g. "Hold 100 shares at 2850").
     * @return Analytical report synthesized with real-time tool executions.
     */
    @SystemMessage("""
            You are a Senior Quantitative Analyst at an institutional investment firm.
            Inspect the specified equity symbol using live tools and evaluate its Trend-Regression-Momentum setup.
            Tag all outputs with [OBSERVED], [CALCULATED], [ASSUMED].
            """)
    @UserMessage("""
            Perform a complete TRM quantitative evaluation for equity symbol: {{symbol}}.
            Trader Position & Portfolio Context: {{userContext}}
            Verify 100-period regression bounds, circuit limits, and binary event risk before giving your directive.
            """)
    String analyzeWithPromptTemplate(@V("symbol") String symbol, @V("userContext") String userContext);

    /**
     * Executes stateful multi-turn conversational advisory using session-isolated memory.
     * 
     * Annotation: @MemoryId
     * Binds this session to a unique conversation history maintained by the ChatMemoryProvider.
     * Real-Time Scenario: A trader asks follow-up questions ("What did we say about the stop loss earlier?")
     * and the AI agent recalls previous dialogue turns specifically for this sessionId.
     * 
     * @param sessionId Unique identifier for the conversation session (e.g. "desk-1-trader-john").
     * @param userMessage The new follow-up message from the user.
     * @return Contextual conversational response preserving dialogue continuity.
     */
    @SystemMessage("""
            You are a conversational algorithmic trading assistant.
            Maintain conversational context across questions. When asked follow-up questions about previously discussed
            stocks, reference prior calculations while confirming live data via tools if market status may have changed.
            """)
    String conversationalAdvisory(@MemoryId String sessionId, @UserMessage String userMessage);

    /**
     * Extracts strongly-typed structured JSON directly into a Java Record without regex parsing.
     * 
     * Return Type: {@link TrmStructuredAdvisory}
     * LangChain4j inspects the Record fields, prompts the model for conforming JSON, and deserializes it automatically.
     * 
     * @param symbol The NSE ticker symbol.
     * @param userContext Optional user scenario or trade constraints.
     * @return Fully populated, strongly-typed TrmStructuredAdvisory Record.
     */
    @SystemMessage("""
            You are a quantitative trading system parser.
            You MUST evaluate the requested symbol using available market data tools.
            Extract exact numeric levels for Last Traded Price, TRM Stop-Loss, Target 1, Target 2, and Setup Score.
            Verify corporate filings to determine if the 48-Hour Binary Event Embargo is CLEAR.
            Output your assessment strictly adhering to the requested JSON schema.
            """)
    @UserMessage("""
            Analyze ticker symbol {{symbol}} for structured quantitative extraction.
            User Scenario: {{userContext}}
            """)
    TrmStructuredAdvisory analyzeStockStructured(@V("symbol") String symbol, @V("userContext") String userContext);

    /**
     * Executes analysis and returns a {@link Result} wrapper containing token usage and execution metadata.
     * 
     * Return Type: Result<String>
     * Provides access to TokenUsage (input tokens, output tokens, total tokens), finish reason, and source metadata.
     * 
     * @param symbol The equity ticker symbol.
     * @return Result object containing the markdown analysis and detailed token metrics.
     */
    @SystemMessage("""
            You are a low-latency equity screening agent. Provide a fast 3-bullet executive summary
            of the stock's trend, regression position, and risk state.
            """)
    @UserMessage("Provide rapid TRM status overview for {{symbol}}.")
    Result<String> analyzeWithResultMetadata(@V("symbol") String symbol);

    /**
     * Performs specialized regulatory and catalyst surveillance using announcement tools.
     * 
     * @param symbol The NSE ticker symbol to audit.
     * @return Clear regulatory risk assessment report.
     */
    @SystemMessage("""
            You are a regulatory surveillance auditor for NSE equities.
            Inspect corporate filings using the getCorporateAnnouncements tool for the given symbol.
            Audit whether any 48-Hour Binary Event Embargo is triggered (such as board meetings for financial results,
            dividend declarations, stock splits, or regulatory investigations).
            Output a concise Risk Clearance Report.
            """)
    String auditCatalystRiskWithTools(@UserMessage String symbol);
}
