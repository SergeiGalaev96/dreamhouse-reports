package com.jasper.jasper_service;

import java.time.Duration;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.*;
import org.springframework.http.client.SimpleClientHttpRequestFactory;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.client.RestClientResponseException;
import org.springframework.web.client.ResourceAccessException;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriComponentsBuilder;

/** Exposes the API-generated report through the common report service. */
@RestController
@CrossOrigin(origins = "*", allowedHeaders = "*",
        exposedHeaders = {HttpHeaders.CONTENT_TYPE, HttpHeaders.CONTENT_DISPOSITION})
public class EstimateChangesController {
    private final RestTemplate client;
    private final String apiBaseUrl;

    public EstimateChangesController(@Value("${reports.api-base-url}") String apiBaseUrl) {
        this.apiBaseUrl = apiBaseUrl;
        SimpleClientHttpRequestFactory factory = new SimpleClientHttpRequestFactory();
        factory.setConnectTimeout(Duration.ofSeconds(10));
        factory.setReadTimeout(Duration.ofSeconds(110));
        this.client = new RestTemplate(factory);
    }

    @GetMapping({"/report/estimate-changes", "/reports/estimate-changes"})
    public ResponseEntity<byte[]> report(
            @RequestParam(name = "material_estimate_id") Long estimateId,
            @RequestParam(defaultValue = "html") String format,
            @RequestParam(name = "date_from", required = false) String dateFrom,
            @RequestParam(name = "date_to", required = false) String dateTo,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorization) {
        var url = UriComponentsBuilder.fromUriString(apiBaseUrl)
                .path("/api/reports/estimate-changes")
                .queryParam("material_estimate_id", estimateId)
                .queryParam("format", format);
        if (dateFrom != null && !dateFrom.isBlank()) url.queryParam("date_from", dateFrom);
        if (dateTo != null && !dateTo.isBlank()) url.queryParam("date_to", dateTo);
        HttpHeaders headers = new HttpHeaders();
        if (authorization != null) headers.set(HttpHeaders.AUTHORIZATION, authorization);
        try {
            var response = client.exchange(url.build().encode().toUri(), HttpMethod.GET,
                    new HttpEntity<Void>(headers), byte[].class);
            return forward(response.getStatusCode(), response.getHeaders(), response.getBody());
        } catch (RestClientResponseException error) {
            return forward(error.getStatusCode(), error.getResponseHeaders(), error.getResponseBodyAsByteArray());
        } catch (ResourceAccessException error) {
            return ResponseEntity.status(502).contentType(MediaType.APPLICATION_JSON)
                    .body("{\"success\":false,\"message\":\"Report API is unavailable\"}".getBytes(java.nio.charset.StandardCharsets.UTF_8));
        }
    }

    private ResponseEntity<byte[]> forward(HttpStatusCode status, HttpHeaders source, byte[] body) {
        HttpHeaders headers = new HttpHeaders();
        if (source != null) {
            for (String name : new String[]{HttpHeaders.CONTENT_TYPE, HttpHeaders.CONTENT_DISPOSITION}) {
                if (source.get(name) != null) headers.put(name, source.get(name));
            }
        }
        headers.setCacheControl("no-store");
        return new ResponseEntity<>(body, headers, status);
    }
}
