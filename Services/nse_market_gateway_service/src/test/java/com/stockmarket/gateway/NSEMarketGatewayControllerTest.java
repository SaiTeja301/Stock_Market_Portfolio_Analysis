package com.stockmarket.gateway;

import com.stockmarket.gateway.controller.NSEMarketGatewayController;
import com.stockmarket.gateway.dto.SymbolQuoteRecord;
import com.stockmarket.gateway.service.NSEMarketGatewayService;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.MockMvc;

import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

@WebMvcTest(NSEMarketGatewayController.class)
class NSEMarketGatewayControllerTest {

    @Autowired
    private MockMvc mockMvc;

    @MockBean
    private NSEMarketGatewayService gatewayService;

    @Test
    @DisplayName("GET /api/v1/market/quote/{symbol} should return 200 OK with quote")
    void testGetQuoteEndpoint() throws Exception {
        SymbolQuoteRecord quote = new SymbolQuoteRecord(
                "BHARTIARTL", "EQ", 1845.50, 15.20, 0.83, 1830.30,
                1835.00, 1845.50, 1840.10, 1655.00, 2022.00, 1900.00, 1100.00, null
        );

        when(gatewayService.getSymbolQuote("BHARTIARTL")).thenReturn(quote);

        mockMvc.perform(get("/api/v1/market/quote/BHARTIARTL")
                        .accept(MediaType.APPLICATION_JSON))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.symbol").value("BHARTIARTL"))
                .andExpect(jsonPath("$.lastPrice").value(1845.50))
                .andExpect(jsonPath("$.vwap").value(1840.10));
    }
}
