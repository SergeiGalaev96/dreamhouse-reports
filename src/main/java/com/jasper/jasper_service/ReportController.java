package com.jasper.jasper_service;

import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestHeader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Set;

@RestController
@CrossOrigin(
        origins = "*",
        allowedHeaders = "*",
        exposedHeaders = {HttpHeaders.CONTENT_DISPOSITION, HttpHeaders.CONTENT_TYPE}
)
@RequestMapping("/report")
public class ReportController {

    private static final Set<String> SUPPORTED_FORMATS = Set.of("pdf", "html", "xlsx", "docx");

    private final ReportService reportService;

    public ReportController(ReportService reportService) {
        this.reportService = reportService;
    }

    @GetMapping("/workPerformed/{id}")
    public ResponseEntity<byte[]> generate(
            @PathVariable String id,
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
        ) throws Exception {

        String normalizedFormat = normalizeFormat(format);
        ReportService.ReportData reportData = reportService.getWorkPerformedData(id, authorizationHeader);
        byte[] file = reportService.generateWorkPerformed(reportData, id, normalizedFormat);
        String filename = reportService.buildWorkPerformedFilename(reportData, id, normalizedFormat);

        String contentType = switch (normalizedFormat) {
            case "html" -> MediaType.TEXT_HTML_VALUE;
            case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            default -> MediaType.APPLICATION_PDF_VALUE;
        };

        String disposition = normalizedFormat.equals("html")
                ? "inline"
                : "attachment";

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_TYPE, contentType)
                .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                .body(file);
    }

    @GetMapping("/form29")
    public ResponseEntity<byte[]> generateForm29(
            @RequestParam Integer blockId,
            @RequestParam String dateFrom,
            @RequestParam String dateTo,
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            ReportService.ReportData reportData = reportService.getForm29Data(blockId, dateFrom, dateTo, authorizationHeader);
            byte[] file = reportService.generateForm29(reportData, normalizedFormat);
            String filename = reportService.buildForm29Filename(reportData, normalizedFormat);

            String contentType = switch (normalizedFormat) {
                case "html" -> MediaType.TEXT_HTML_VALUE + "; charset=UTF-8";
                case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default -> MediaType.APPLICATION_PDF_VALUE;
            };

            String disposition = normalizedFormat.equals("html")
                    ? "inline"
                    : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Form29 generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/form2")
    public ResponseEntity<byte[]> generateForm2(
            @RequestParam Integer blockId,
            @RequestParam String dateFrom,
            @RequestParam String dateTo,
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            ReportService.ReportData reportData = reportService.getForm2Data(blockId, dateFrom, dateTo, authorizationHeader);
            byte[] file = reportService.generateForm2(reportData, normalizedFormat);
            String filename = reportService.buildForm2Filename(reportData, normalizedFormat);

            String contentType = switch (normalizedFormat) {
                case "html" -> MediaType.TEXT_HTML_VALUE + "; charset=UTF-8";
                case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default -> MediaType.APPLICATION_PDF_VALUE;
            };

            String disposition = normalizedFormat.equals("html")
                    ? "inline"
                    : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Form2 generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/form19")
    public ResponseEntity<byte[]> generateForm19(
            @RequestParam Integer projectId,
            @RequestParam String dateFrom,
            @RequestParam String dateTo,
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            ReportService.ReportData reportData = reportService.getForm19Data(projectId, dateFrom, dateTo, authorizationHeader);
            byte[] file = reportService.generateForm19(reportData, normalizedFormat);
            String filename = reportService.buildForm19Filename(reportData, normalizedFormat);

            String contentType = switch (normalizedFormat) {
                case "html" -> MediaType.TEXT_HTML_VALUE + "; charset=UTF-8";
                case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default -> MediaType.APPLICATION_PDF_VALUE;
            };

            String disposition = normalizedFormat.equals("html")
                    ? "inline"
                    : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Form19 generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/mbp-write-off")
    public ResponseEntity<byte[]> generateMbpWriteOff(
            @RequestParam Integer projectId,
            @RequestParam String dateFrom,
            @RequestParam String dateTo,
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            ReportService.ReportData reportData = reportService.getMbpWriteOffData(projectId, dateFrom, dateTo, authorizationHeader);
            byte[] file = reportService.generateMbpWriteOff(reportData, normalizedFormat);
            String filename = reportService.buildMbpWriteOffFilename(reportData, normalizedFormat);

            String contentType = switch (normalizedFormat) {
                case "html" -> MediaType.TEXT_HTML_VALUE + "; charset=UTF-8";
                case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default -> MediaType.APPLICATION_PDF_VALUE;
            };

            String disposition = normalizedFormat.equals("html")
                    ? "inline"
                    : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "MBP write-off generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/projects-overview")
    public ResponseEntity<byte[]> generateProjectsOverview(
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            ReportService.ReportData reportData = reportService.getProjectsOverviewData(authorizationHeader);
            byte[] file = reportService.generateProjectsOverview(reportData, normalizedFormat);
            String filename = reportService.buildProjectsOverviewFilename(reportData, normalizedFormat);

            String contentType = switch (normalizedFormat) {
                case "html" -> MediaType.TEXT_HTML_VALUE + "; charset=UTF-8";
                case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default -> MediaType.APPLICATION_PDF_VALUE;
            };

            String disposition = normalizedFormat.equals("html")
                    ? "inline"
                    : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Projects overview generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/estimate-stage")
    public ResponseEntity<byte[]> generateEstimateStage(
            @RequestParam Integer blockId,
            @RequestParam(defaultValue = "xlsx") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = format == null ? "xlsx" : format.toLowerCase();
            if (!"xlsx".equals(normalizedFormat) && !"html".equals(normalizedFormat)) {
                normalizedFormat = "xlsx";
            }

            ReportService.ReportData reportData = reportService.getEstimateStageData(blockId, authorizationHeader);
            byte[] file = reportService.generateEstimateStage(reportData, normalizedFormat);
            String filename = reportService.buildEstimateStageFilename(reportData, normalizedFormat);

            String contentType = "html".equals(normalizedFormat)
                    ? MediaType.TEXT_HTML_VALUE + "; charset=UTF-8"
                    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            String disposition = "html".equals(normalizedFormat) ? "inline" : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Estimate stage generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/paymentCashOrder/{id}")
    public ResponseEntity<byte[]> generatePaymentCashOrder(
            @PathVariable String id,
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            ReportService.ReportData reportData = reportService.getPaymentCashOrderData(id, authorizationHeader);
            byte[] file = reportService.generatePaymentCashOrder(reportData, normalizedFormat);
            String filename = reportService.buildPaymentCashOrderFilename(reportData, normalizedFormat);

            String contentType = switch (normalizedFormat) {
                case "html" -> MediaType.TEXT_HTML_VALUE + "; charset=UTF-8";
                case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default -> MediaType.APPLICATION_PDF_VALUE;
            };

            String disposition = normalizedFormat.equals("html")
                    ? "inline"
                    : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Payment cash order generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/payments-cash-book")
    public ResponseEntity<byte[]> generatePaymentsCashBook(
            @RequestParam Integer projectId,
            @RequestParam(required = false) Integer blockId,
            @RequestParam(required = false) String dateFrom,
            @RequestParam(required = false) String dateTo,
            @RequestParam(defaultValue = "xlsx") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            if (!"html".equals(normalizedFormat) && !"xlsx".equals(normalizedFormat)) {
                normalizedFormat = "xlsx";
            }

            ReportService.ReportData reportData = reportService.getPaymentsCashBookData(projectId, blockId, dateFrom, dateTo, authorizationHeader);
            byte[] file = reportService.generatePaymentsCashBook(reportData, normalizedFormat);
            String filename = reportService.buildPaymentsCashBookFilename(reportData);
            String contentType = "html".equals(normalizedFormat)
                    ? MediaType.TEXT_HTML_VALUE + "; charset=UTF-8"
                    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            String disposition = "html".equals(normalizedFormat) ? "inline" : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Payments cash book generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/sales-summary")
    public ResponseEntity<byte[]> generateSalesSummary(
            @RequestParam(required = false) Integer projectId,
            @RequestParam(required = false) String projectIds,
            @RequestParam(required = false) Integer blockId,
            @RequestParam(required = false) String dateFrom,
            @RequestParam(required = false) String dateTo,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            ReportService.ReportData reportData = reportService.getSalesSummaryData(projectId, projectIds, blockId, dateFrom, dateTo, authorizationHeader);
            byte[] file = reportService.generateSalesSummary(reportData);
            String filename = reportService.buildSalesSummaryFilename(reportData);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition("attachment", filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Sales summary generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/sales-payment-schedule")
    public ResponseEntity<byte[]> generateSalesPaymentSchedule(
            @RequestParam(required = false) Integer projectId,
            @RequestParam(required = false) Integer blockId,
            @RequestParam(required = false) Integer dealId,
            @RequestParam(required = false) String dateFrom,
            @RequestParam(required = false) String dateTo,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            ReportService.ReportData reportData = reportService.getSalesPaymentScheduleData(projectId, blockId, dealId, dateFrom, dateTo, authorizationHeader);
            byte[] file = reportService.generateSalesPaymentSchedule(reportData);
            String filename = reportService.buildSalesPaymentScheduleFilename(reportData);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition("attachment", filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Sales payment schedule generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/sales-monthly-plan")
    public ResponseEntity<byte[]> generateSalesMonthlyPlan(
            @RequestParam Integer projectId,
            @RequestParam(required = false) Integer blockId,
            @RequestParam(defaultValue = "xlsx") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = "html".equalsIgnoreCase(format) ? "html" : "xlsx";
            ReportService.ReportData reportData = reportService.getSalesMonthlyPlanData(projectId, blockId, authorizationHeader);
            byte[] file = reportService.generateSalesMonthlyPlan(reportData, normalizedFormat);
            String filename = reportService.buildSalesMonthlyPlanFilename(reportData, normalizedFormat);
            String contentType = "html".equals(normalizedFormat)
                    ? MediaType.TEXT_HTML_VALUE + "; charset=UTF-8"
                    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition("html".equals(normalizedFormat) ? "inline" : "attachment", filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Sales monthly plan generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/schedule")
    public ResponseEntity<byte[]> generateSchedule(
            @RequestParam Integer projectId,
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            ReportService.ReportData reportData = reportService.getScheduleData(projectId, authorizationHeader);
            byte[] file = reportService.generateSchedule(reportData, normalizedFormat);
            String filename = reportService.buildScheduleFilename(reportData, normalizedFormat);

            String contentType = switch (normalizedFormat) {
                case "html" -> MediaType.TEXT_HTML_VALUE + "; charset=UTF-8";
                case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default -> MediaType.APPLICATION_PDF_VALUE;
            };

            String disposition = normalizedFormat.equals("html")
                    ? "inline"
                    : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Schedule generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    @GetMapping("/material-schedule")
    public ResponseEntity<byte[]> generateMaterialSchedule(
            @RequestParam Integer projectId,
            @RequestParam(defaultValue = "pdf") String format,
            @RequestHeader(value = HttpHeaders.AUTHORIZATION, required = false) String authorizationHeader
    ) throws Exception {
        try {
            String normalizedFormat = normalizeFormat(format);
            ReportService.ReportData reportData = reportService.getMaterialScheduleData(projectId, authorizationHeader);
            byte[] file = reportService.generateMaterialSchedule(reportData, normalizedFormat);
            String filename = reportService.buildScheduleFilename(reportData, normalizedFormat);

            String contentType = switch (normalizedFormat) {
                case "html" -> MediaType.TEXT_HTML_VALUE + "; charset=UTF-8";
                case "xlsx" -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case "docx" -> "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                default -> MediaType.APPLICATION_PDF_VALUE;
            };

            String disposition = normalizedFormat.equals("html")
                    ? "inline"
                    : "attachment";

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_TYPE, contentType)
                    .header(HttpHeaders.CONTENT_DISPOSITION, contentDisposition(disposition, filename))
                    .body(file);
        } catch (Exception e) {
            String message = e.getClass().getSimpleName() + ": " + (e.getMessage() == null ? "Material schedule generation failed" : e.getMessage());
            return ResponseEntity.internalServerError()
                    .header(HttpHeaders.CONTENT_TYPE, MediaType.TEXT_PLAIN_VALUE + "; charset=UTF-8")
                    .body(message.getBytes(StandardCharsets.UTF_8));
        }
    }

    private String normalizeFormat(String format) {
        String normalized = format == null ? "pdf" : format.toLowerCase();
        if (!SUPPORTED_FORMATS.contains(normalized)) {
            return "pdf";
        }
        return normalized;
    }

    private String contentDisposition(String disposition, String filename) {
        String encoded = URLEncoder.encode(filename, StandardCharsets.UTF_8).replace("+", "%20");
        return disposition + "; filename=\"report." + extension(filename) + "\"; filename*=UTF-8''" + encoded;
    }

    private String extension(String filename) {
        int dotIndex = filename.lastIndexOf('.');
        if (dotIndex < 0 || dotIndex == filename.length() - 1) {
            return "pdf";
        }
        return filename.substring(dotIndex + 1);
    }
}
