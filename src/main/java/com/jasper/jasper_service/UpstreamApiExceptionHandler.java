package com.jasper.jasper_service;

import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.client.RestClientResponseException;

import java.nio.charset.StandardCharsets;

@RestControllerAdvice
public class UpstreamApiExceptionHandler {

    @ExceptionHandler(RestClientResponseException.class)
    public ResponseEntity<byte[]> handleUpstreamResponse(RestClientResponseException exception) {
        byte[] body = exception.getResponseBodyAsByteArray();
        if (body.length == 0) {
            body = exception.getStatusText().getBytes(StandardCharsets.UTF_8);
        }

        MediaType contentType = exception.getResponseHeaders() == null
                ? null
                : exception.getResponseHeaders().getContentType();

        return ResponseEntity.status(exception.getStatusCode())
                .header(
                        HttpHeaders.CONTENT_TYPE,
                        contentType == null
                                ? MediaType.APPLICATION_JSON_VALUE + "; charset=UTF-8"
                                : contentType.toString()
                )
                .body(body);
    }
}
