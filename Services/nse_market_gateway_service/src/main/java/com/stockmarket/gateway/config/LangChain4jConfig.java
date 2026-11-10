package com.stockmarket.gateway.config;

import dev.langchain4j.memory.chat.ChatMemoryProvider;
import dev.langchain4j.memory.chat.MessageWindowChatMemory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * ====================================================================================================
 *                               LANGCHAIN4J ENTERPRISE CONFIGURATION
 * ====================================================================================================
 * 
 * Architectural Role:
 * Configures enterprise infrastructure beans for LangChain4j integration within the Spring Boot container.
 * Specifically registers a centralized {@link ChatMemoryProvider} to enable stateful multi-turn conversations
 * across declarative {@code @AiService} interfaces using the {@code @MemoryId} annotation.
 * 
 * Key LangChain4j Core Concepts Configured Here:
 * 1. ChatMemory:
 *    Maintains the conversational context (SystemMessage, UserMessage, AiMessage, ToolExecutionResultMessage)
 *    between the client and the LLM. Without chat memory, LLMs are completely stateless.
 * 
 * 2. ChatMemoryProvider:
 *    A factory functional interface: {@code (Object memoryId) -> ChatMemory}.
 *    When an {@code @AiService} method annotated with {@code @MemoryId} is invoked, LangChain4j uses this
 *    provider to retrieve or instantiate an isolated, dedicated memory instance for that specific session/user.
 * 
 * 3. MessageWindowChatMemory:
 *    A sliding-window memory eviction strategy that retains the N most recent messages (e.g. 10 messages).
 *    Prevents context window overflow and unbounded memory consumption in high-concurrency environments.
 * 
 * Real-Time Production Scenario:
 * In institutional stock advisory, multiple portfolio managers or trading desks query the same system
 * concurrently. Using {@code @MemoryId("trader-user-101")} or {@code @MemoryId("session-RELIANCE")} ensures
 * that trader A discussing Reliance never cross-contaminates trader B discussing TCS, while preserving
 * full multi-turn follow-up question capabilities within their respective trading desks.
 */
@Slf4j
@Configuration
public class LangChain4jConfig {

    /**
     * Configures the default {@link ChatMemoryProvider} bean for LangChain4j.
     * 
     * Annotation: @Bean
     * Role: Registers the return object as a Spring-managed singleton bean in the ApplicationContext.
     * Mechanics: LangChain4j's auto-configuration (AiServicesAutoConfig) automatically scans the
     *            ApplicationContext for a ChatMemoryProvider bean and binds it to all @AiService proxies.
     * 
     * @return ChatMemoryProvider configured with a 10-message sliding window per unique memoryId.
     */
    @Bean
    public ChatMemoryProvider chatMemoryProvider() {
        log.info("[LangChain4jConfig] Initializing MessageWindowChatMemoryProvider (Sliding Window: 10 messages per MemoryId)");
        return memoryId -> {
            log.debug("[LangChain4j ChatMemory] Allocating or retrieving ChatMemory window for memoryId: {}", memoryId);
            return MessageWindowChatMemory.builder()
                    .id(memoryId)
                    .maxMessages(10)
                    .build();
        };
    }
}
