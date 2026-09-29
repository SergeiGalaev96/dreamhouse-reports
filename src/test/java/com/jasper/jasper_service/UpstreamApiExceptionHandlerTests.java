package com.jasper.jasper_service;

import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.client.HttpClientErrorException;

import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;

class UpstreamApiExceptionHandlerTests {

    @Test
    void preservesUnauthorizedResponseFromApi() {
        byte[] body = "{\"success\":false,\"code\":\"ACCESS_TOKEN_EXPIRED\"}"
                .getBytes(StandardCharsets.UTF_8);
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_JSON);
        HttpClientErrorException exception = HttpClientErrorException.create(
                HttpStatus.UNAUTHORIZED,
                "Unauthorized",
                headers,
                body,
                StandardCharsets.UTF_8
        );

        ResponseEntity<byte[]> response = new UpstreamApiExceptionHandler()
                .handleUpstreamResponse(exception);

        assertEquals(HttpStatus.UNAUTHORIZED, response.getStatusCode());
        assertEquals(MediaType.APPLICATION_JSON, response.getHeaders().getContentType());
        assertArrayEquals(body, response.getBody());
    }
}
