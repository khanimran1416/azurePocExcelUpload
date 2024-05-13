package com.hdfcbank.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.reactive.function.client.WebClient;

@Configuration
public class WebConfig {

    @Bean
    public WebClient webClient() {

        WebClient client = WebClient.builder()
                .baseUrl("http://apigw-hdfcpoc.centralindia.cloudapp.azure.com")
                .defaultCookie("cookieKey", "cookieValue")
                .defaultHeaders(httpHeaders -> {
                    httpHeaders.set(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_JSON_VALUE);
                    httpHeaders.set("x-requesting-user","aftab");
                })
                .build();
        return client;
    }
}
