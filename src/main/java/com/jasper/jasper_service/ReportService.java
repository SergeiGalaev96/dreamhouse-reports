package com.jasper.jasper_service;

import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JREmptyDataSource;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.data.JRMapCollectionDataSource;
import net.sf.jasperreports.engine.export.HtmlExporter;
import net.sf.jasperreports.engine.export.ooxml.JRDocxExporter;
import net.sf.jasperreports.engine.export.ooxml.JRXlsxExporter;
import net.sf.jasperreports.export.SimpleExporterInput;
import net.sf.jasperreports.export.SimpleHtmlExporterOutput;
import net.sf.jasperreports.export.SimpleOutputStreamExporterOutput;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.FontMetrics;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Collection;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.nio.charset.StandardCharsets;

@Service
public class ReportService {

    private final RestTemplate restTemplate = new RestTemplate();

    @Value("${reports.api-base-url}")
    private String apiBaseUrl;

    public ReportData getWorkPerformedData(String id, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/workPerformed/" + id;

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.GET,
                new HttpEntity<>(headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Report API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Report API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Report API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getForm29Data(Integer blockId, String dateFrom, String dateTo, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/form29";

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        Map<String, Object> payload = new HashMap<>();
        payload.put("block_id", blockId);
        payload.put("date_from", dateFrom);
        payload.put("date_to", dateTo);

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(payload, headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Form29 API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Form29 API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Form29 API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getForm2Data(Integer blockId, String dateFrom, String dateTo, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/form2";

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        Map<String, Object> payload = new HashMap<>();
        payload.put("block_id", blockId);
        payload.put("date_from", dateFrom);
        payload.put("date_to", dateTo);

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(payload, headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Form2 API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Form2 API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Form2 API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getMaterialRequestData(String id, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/materialRequest/" + id;

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.GET,
                new HttpEntity<>(headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Material request report API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        Map header = data == null ? null : (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Material request report API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getForm19Data(Integer projectId, String dateFrom, String dateTo, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/form19";

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        Map<String, Object> payload = new HashMap<>();
        payload.put("project_id", projectId);
        payload.put("date_from", dateFrom);
        payload.put("date_to", dateTo);

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(payload, headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Form19 API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Form19 API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Form19 API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getMbpWriteOffData(Integer projectId, String dateFrom, String dateTo, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/mbp-write-off";

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        Map<String, Object> payload = new HashMap<>();
        payload.put("project_id", projectId);
        payload.put("date_from", dateFrom);
        payload.put("date_to", dateTo);

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(payload, headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("MBP write-off API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("MBP write-off API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("MBP write-off API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getProjectsOverviewData(String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/projects-overview";

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(Map.of(), headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Projects overview API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Projects overview API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Projects overview API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getEstimateStageData(Integer blockId, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/estimate-stage";

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        Map<String, Object> payload = new HashMap<>();
        payload.put("block_id", blockId);

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(payload, headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Estimate stage API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Estimate stage API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Estimate stage API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getScheduleData(Integer projectId, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/schedule";

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        Map<String, Object> payload = new HashMap<>();
        payload.put("project_id", projectId);

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(payload, headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Schedule API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Schedule API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Schedule API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getMaterialScheduleData(Integer projectId, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/material-schedule";

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        Map<String, Object> payload = new HashMap<>();
        payload.put("project_id", projectId);

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(payload, headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Material schedule API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Material schedule API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Material schedule API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getPaymentCashOrderData(String id, String authorizationHeader) {
        String url = apiBaseUrl + "/api/reports/paymentCashOrder/" + id;

        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.GET,
                new HttpEntity<>(headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException("Payment cash order API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException("Payment cash order API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException("Payment cash order API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public ReportData getPaymentsCashBookData(
            Integer projectId,
            Integer blockId,
            String dateFrom,
            String dateTo,
            String authorizationHeader
    ) {
        String url = apiBaseUrl + "/api/reports/payments/cash-book";
        Map<String, Object> payload = new HashMap<>();
        if (projectId != null) {
            payload.put("project_id", projectId);
        }
        if (blockId != null) {
            payload.put("block_id", blockId);
        }
        if (dateFrom != null && !dateFrom.isBlank()) {
            payload.put("date_from", dateFrom);
        }
        if (dateTo != null && !dateTo.isBlank()) {
            payload.put("date_to", dateTo);
        }
        return getReportDataByPost(url, payload, authorizationHeader, "Payments cash book");
    }

    public ReportData getSalesSummaryData(
            Integer projectId,
            String projectIds,
            Integer blockId,
            String dateFrom,
            String dateTo,
            String authorizationHeader
    ) {
        String url = apiBaseUrl + "/api/reports/sales/summary";
        Map<String, Object> payload = new HashMap<>();
        if (projectId != null) {
            payload.put("project_id", projectId);
        }
        if (projectIds != null && !projectIds.isBlank()) {
            payload.put("project_ids", projectIds);
        }
        if (blockId != null) {
            payload.put("block_id", blockId);
        }
        if (dateFrom != null && !dateFrom.isBlank()) {
            payload.put("date_from", dateFrom);
        }
        if (dateTo != null && !dateTo.isBlank()) {
            payload.put("date_to", dateTo);
        }
        return getReportDataByPost(url, payload, authorizationHeader, "Sales summary");
    }

    public ReportData getSalesPaymentScheduleData(
            Integer projectId,
            Integer blockId,
            Integer dealId,
            String dateFrom,
            String dateTo,
            String authorizationHeader
    ) {
        String url = apiBaseUrl + "/api/reports/sales/payment-schedule";
        Map<String, Object> payload = new HashMap<>();
        if (projectId != null) {
            payload.put("project_id", projectId);
        }
        if (blockId != null) {
            payload.put("block_id", blockId);
        }
        if (dealId != null) {
            payload.put("deal_id", dealId);
        }
        if (dateFrom != null && !dateFrom.isBlank()) {
            payload.put("date_from", dateFrom);
        }
        if (dateTo != null && !dateTo.isBlank()) {
            payload.put("date_to", dateTo);
        }
        return getReportDataByPost(url, payload, authorizationHeader, "Sales payment schedule");
    }

    public ReportData getSalesMonthlyPlanData(
            Integer projectId,
            Integer blockId,
            String authorizationHeader
    ) {
        String url = apiBaseUrl + "/api/reports/sales/monthly-plan";
        Map<String, Object> payload = new HashMap<>();
        payload.put("project_id", projectId);
        if (blockId != null) {
            payload.put("block_id", blockId);
        }
        return getReportDataByPost(url, payload, authorizationHeader, "Sales monthly plan");
    }

    private ReportData getReportDataByPost(
            String url,
            Map<String, Object> payload,
            String authorizationHeader,
            String reportName
    ) {
        HttpHeaders headers = new HttpHeaders();
        if (authorizationHeader != null && !authorizationHeader.isBlank()) {
            headers.set(HttpHeaders.AUTHORIZATION, authorizationHeader);
        }

        ResponseEntity<Map> apiResponse = restTemplate.exchange(
                url,
                HttpMethod.POST,
                new HttpEntity<>(payload, headers),
                Map.class
        );

        Map response = apiResponse.getBody();
        if (response == null || !Boolean.TRUE.equals(response.get("success"))) {
            throw new IllegalStateException(reportName + " API returned an unsuccessful response");
        }

        Map data = (Map) response.get("data");
        if (data == null) {
            throw new IllegalStateException(reportName + " API response does not contain data");
        }

        Map header = (Map) data.get("header");
        if (header == null) {
            throw new IllegalStateException(reportName + " API response does not contain header");
        }

        return new ReportData(data, header);
    }

    public byte[] generateWorkPerformed(String id, String format, String authorizationHeader) throws Exception {
        return generateWorkPerformed(getWorkPerformedData(id, authorizationHeader), id, format);
    }

    public byte[] generateWorkPerformed(ReportData reportData, String id, String format) throws Exception {
        InputStream jrxmlStream =
                new ClassPathResource("reports/WorkPerformedReport.jrxml").getInputStream();

        JasperReport report = JasperCompileManager.compileReport(jrxmlStream);

        Map data = reportData.data();
        Map header = reportData.header();

        List<Map<String, Object>> itemsFromApi =
                (List<Map<String, Object>>) data.get("items");

        List<Map<String, Object>> items = new ArrayList<>();

        for (Map<String, Object> item : itemsFromApi) {
            Map<String, Object> row = new HashMap<>();

            row.put("service_name", firstNonNull(item.get("service_name"), item.get("name")));
            row.put("stage_name", item.get("stage_name"));
            row.put("subsection_name", item.get("subsection_name"));
            row.put("unit_name", item.get("unit_name"));
            row.put("quantity", toDouble(item.get("quantity")));
            row.put("price", toDouble(item.getOrDefault("price_converted", item.get("price"))));
            row.put("total", toDouble(item.get("total")));

            items.add(row);
        }

        Collection<Map<String, ?>> collection = new ArrayList<>(items);
        JRMapCollectionDataSource dataSource = new JRMapCollectionDataSource(collection);

        Map<String, Object> params = new HashMap<>();
        double total = toDouble(data.get("total"));
        double advancePayment = toDouble(firstNonNull(data.get("advance_payment"), header.get("advance_payment")));
        double remainingAmount = total - advancePayment;
        params.put("header", header);
        params.put("total", total);
        params.put("advancePayment", advancePayment);
        params.put("remainingAmount", remainingAmount);
        params.put("totalInWords", amountToWords(total));
        params.put("reportId", firstNonBlank(header.get("id"), id));
        params.put("createdAt", formatCreatedAt(header.get("created_at")));
        params.put("projectName", stringValue(header.get("project_name")));
        params.put("blockName", stringValue(header.get("block_name")));
        params.put("statusName", stringValue(header.get("status_name")));
        params.put("performerName", stringValue(header.get("performed_person_name")));
        params.put("foremanName", stringValue(firstNonNull(
                header.get("foreman_name"),
                header.get("assigned_foreman_name")
        )));
        params.put("planningEngineerName", stringValue(firstNonNull(
                header.get("planning_engineer_name"),
                header.get("assigned_planning_engineer_name")
        )));
        params.put("mainEngineerName", stringValue(firstNonNull(
                header.get("main_engineer_name"),
                header.get("assigned_main_engineer_name")
        )));
        params.put("performerSignatureName", formatSignatureName(header.get("performed_person_name")));
        params.put("foremanSignatureName", formatSignatureName(firstNonNull(
                header.get("foreman_name"),
                header.get("assigned_foreman_name")
        )));
        params.put("planningEngineerSignatureName", formatSignatureName(firstNonNull(
                header.get("planning_engineer_name"),
                header.get("assigned_planning_engineer_name")
        )));
        params.put("mainEngineerSignatureName", formatSignatureName(firstNonNull(
                header.get("main_engineer_name"),
                header.get("assigned_main_engineer_name")
        )));

        JasperPrint print = JasperFillManager.fillReport(
                report,
                params,
                dataSource
        );

        return switch (format) {
            case "xlsx" -> exportXlsx(print);
            case "docx" -> exportDocx(print);
            case "html" -> exportHtml(print);
            default -> JasperExportManager.exportReportToPdf(print);
        };
    }

    public byte[] generateMaterialRequest(ReportData reportData) throws Exception {
        Map header = reportData.header();
        List<Map<String, Object>> items = (List<Map<String, Object>>) reportData.data().getOrDefault("items", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Заявка на материалы");
            sheet.setFitToPage(true);
            sheet.setHorizontallyCenter(true);
            sheet.getPrintSetup().setLandscape(false);
            sheet.setAutobreaks(true);
            sheet.setColumnWidth(0, 7 * 256);
            sheet.setColumnWidth(1, 38 * 256);
            sheet.setColumnWidth(2, 10 * 256);
            sheet.setColumnWidth(3, 11 * 256);
            sheet.setColumnWidth(4, 12 * 256);
            sheet.setColumnWidth(5, 14 * 256);
            sheet.setColumnWidth(6, 31 * 256);

            Font boldFont = workbook.createFont();
            boldFont.setBold(true);
            Font normalFont = workbook.createFont();
            normalFont.setFontName("Times New Roman");

            CellStyle textStyle = workbook.createCellStyle();
            textStyle.setFont(normalFont);
            textStyle.setWrapText(true);
            textStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            CellStyle centeredStyle = workbook.createCellStyle();
            centeredStyle.cloneStyleFrom(textStyle);
            centeredStyle.setAlignment(HorizontalAlignment.CENTER);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.cloneStyleFrom(centeredStyle);
            headerStyle.setFont(boldFont);
            headerStyle.setBorderTop(BorderStyle.THIN);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            headerStyle.setBorderLeft(BorderStyle.THIN);
            headerStyle.setBorderRight(BorderStyle.THIN);

            CellStyle tableTextStyle = workbook.createCellStyle();
            tableTextStyle.cloneStyleFrom(textStyle);
            tableTextStyle.setBorderTop(BorderStyle.THIN);
            tableTextStyle.setBorderBottom(BorderStyle.THIN);
            tableTextStyle.setBorderLeft(BorderStyle.THIN);
            tableTextStyle.setBorderRight(BorderStyle.THIN);

            CellStyle tableCenterStyle = workbook.createCellStyle();
            tableCenterStyle.cloneStyleFrom(headerStyle);
            tableCenterStyle.setFont(normalFont);

            int rowIndex = 0;
            Row approvalRow = sheet.createRow(rowIndex++);
            approvalRow.createCell(4).setCellValue("Утверждаю в сумме ____________");
            Row directorRow = sheet.createRow(rowIndex++);
            directorRow.createCell(4).setCellValue("Генеральный директор __________________");

            Row requestRow = sheet.createRow(rowIndex++);
            Cell requestTitle = requestRow.createCell(0);
            requestTitle.setCellValue("Заявка №" + stringValue(header.get("id")));
            requestTitle.setCellStyle(centeredStyle);
            requestTitle.getCellStyle().setFont(boldFont);
            sheet.addMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 6));

            Row objectRow = sheet.createRow(rowIndex++);
            objectRow.createCell(0).setCellValue("По объекту: " + stringValue(header.get("project_name")) + ", " + stringValue(header.get("block_name")));
            sheet.addMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 6));
            Row dateRow = sheet.createRow(rowIndex++);
            dateRow.createCell(0).setCellValue(formatCreatedAt(header.get("created_at")));
            sheet.addMergedRegion(new CellRangeAddress(rowIndex - 1, rowIndex - 1, 0, 6));

            String[] columns = {"№\nп/п", "Наименования", "Ед.\nизм.", "Кол-во", "Цена", "Сумма", "Примечание"};
            Row columnRow = sheet.createRow(rowIndex++);
            columnRow.setHeightInPoints(32);
            for (int column = 0; column < columns.length; column++) {
                Cell cell = columnRow.createCell(column);
                cell.setCellValue(columns[column]);
                cell.setCellStyle(headerStyle);
            }

            int number = 1;
            for (Map<String, Object> item : items) {
                Row row = sheet.createRow(rowIndex++);
                row.setHeightInPoints(30);
                String stageNote = String.join(" — ", List.of(
                        stringValue(item.get("stage_name")),
                        stringValue(item.get("subsection_name"))
                ).stream().filter(value -> !value.isBlank()).toList());
                String note = firstNonBlank(item.get("comment"), stageNote);
                Object[] values = {
                        number++,
                        stringValue(item.get("material_name")),
                        stringValue(item.get("unit_name")),
                        toDouble(item.get("quantity")),
                        toDouble(item.get("price")),
                        toDouble(item.get("total")),
                        note
                };
                for (int column = 0; column < values.length; column++) {
                    Cell cell = row.createCell(column);
                    if (values[column] instanceof Number value) cell.setCellValue(value.doubleValue());
                    else cell.setCellValue(stringValue(values[column]));
                    cell.setCellStyle(column == 1 || column == 6 ? tableTextStyle : tableCenterStyle);
                }
            }

            rowIndex += 2;
            String[][] signatures = {
                    {"Прораб:", stringValue(header.get("foreman_name"))},
                    {"Нач. участка:", stringValue(header.get("site_manager_name"))},
                    {"Снабженец:", stringValue(header.get("purchasing_agent_name"))},
                    {"Инженер ПТО:", stringValue(header.get("planning_engineer_name"))},
                    {"Гл. инженер:", stringValue(header.get("main_engineer_name"))}
            };
            for (String[] signature : signatures) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(signature[0] + " ____________________");
                row.createCell(3).setCellValue(signature[1]);
                sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, 2));
                sheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 3, 6));
            }

            sheet.getPrintSetup().setFitWidth((short) 1);
            sheet.getPrintSetup().setFitHeight((short) 0);
            workbook.write(output);
            return output.toByteArray();
        }
    }

    public byte[] generatePaymentCashOrder(ReportData reportData, String format) throws Exception {
        if ("xlsx".equalsIgnoreCase(format)) {
            return generatePaymentCashOrderXlsx(reportData);
        }

        InputStream jrxmlStream =
                new ClassPathResource("reports/PaymentCashOrderReport.jrxml").getInputStream();

        JasperReport report = JasperCompileManager.compileReport(jrxmlStream);
        Map header = reportData.header();

        double amount = toDouble(header.get("amount"));
        double amountKgs = toDouble(firstNonNull(header.get("amount_kgs"), amount));
        String currency = firstNonBlank(header.get("currency"), "KGS");
        String orderShortName = firstNonBlank(header.get("order_short_name"), "ПКО");

        Map<String, Object> params = new HashMap<>();
        params.put("orderTypeName", firstNonBlank(header.get("order_type_name"), "Кассовый ордер"));
        params.put("orderShortName", orderShortName);
        params.put("orderNumber", firstNonBlank(header.get("order_number"), header.get("id")));
        params.put("orderDate", stringValue(header.get("order_date")));
        params.put("organizationName", firstNonBlank(header.get("organization_name"), "ОсОО \"СК Дрим Хаус\""));
        params.put("projectName", stringValue(header.get("project_name")));
        params.put("blockName", stringValue(header.get("block_name")));
        params.put("articleName", stringValue(header.get("article_name")));
        params.put("title", stringValue(header.get("title")));
        params.put("statusName", stringValue(header.get("status_name")));
        params.put("operationLabel", firstNonBlank(header.get("operation_label"), "Контрагент"));
        params.put("counterpartyName", stringValue(header.get("counterparty_name")));
        params.put("counterpartyInn", stringValue(header.get("counterparty_inn")));
        params.put("entityLabel", stringValue(header.get("entity_label")));
        params.put("basis", stringValue(header.get("basis")));
        params.put("createdByName", stringValue(header.get("created_by_name")));
        params.put("amountLabel", formatSmartNumber(amount) + " " + currency);
        params.put("amountKgsLabel", formatSmartNumber(amountKgs) + " KGS");
        params.put("currencyRateLabel", stringValue(header.get("currency_rate_label")));
        params.put("amountInWords", formatSmartNumber(amount) + " " + currency);

        JasperPrint print = JasperFillManager.fillReport(
                report,
                params,
                new JREmptyDataSource(1)
        );

        return switch (format) {
            case "xlsx" -> exportXlsx(print);
            case "docx" -> exportDocx(print);
            case "html" -> exportHtml(print);
            default -> JasperExportManager.exportReportToPdf(print);
        };
    }

    private byte[] generatePaymentCashOrderXlsx(ReportData reportData) throws Exception {
        Map header = reportData.header();
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(firstNonBlank(header.get("order_short_name"), "Order"));
        sheet.setDisplayGridlines(true);

        for (int i = 0; i < 37; i++) {
            sheet.setColumnWidth(i, 1050);
        }
        for (int i = 16; i <= 23; i++) {
            sheet.setColumnWidth(i, 1250);
        }
        for (int i = 25; i <= 33; i++) {
            sheet.setColumnWidth(i, 1250);
        }
        for (int i = 0; i <= 34; i++) {
            Row row = sheet.createRow(i);
            row.setHeightInPoints(18);
        }
        setRowHeight(sheet, 1, 20);
        setRowHeight(sheet, 2, 20);
        setRowHeight(sheet, 3, 30);
        setRowHeight(sheet, 11, 24);
        setRowHeight(sheet, 13, 24);
        setRowHeight(sheet, 14, 42);
        setRowHeight(sheet, 15, 24);
        setRowHeight(sheet, 16, 24);
        setRowHeight(sheet, 18, 24);
        setRowHeight(sheet, 19, 34);
        setRowHeight(sheet, 21, 34);

        CellStyle normal = createCashOrderStyle(workbook, false, 10, HorizontalAlignment.LEFT, BorderStyle.NONE);
        CellStyle center = createCashOrderStyle(workbook, false, 10, HorizontalAlignment.CENTER, BorderStyle.NONE);
        CellStyle right = createCashOrderStyle(workbook, false, 10, HorizontalAlignment.RIGHT, BorderStyle.NONE);
        CellStyle bold = createCashOrderStyle(workbook, true, 10, HorizontalAlignment.LEFT, BorderStyle.NONE);
        CellStyle title = createCashOrderStyle(workbook, true, 14, HorizontalAlignment.CENTER, BorderStyle.NONE);
        CellStyle orderTitleStyle = createCashOrderStyle(workbook, true, 13, HorizontalAlignment.CENTER, BorderStyle.NONE);
        CellStyle smallCenter = createCashOrderStyle(workbook, false, 8, HorizontalAlignment.CENTER, BorderStyle.NONE);
        CellStyle smallCenterShrink = createCashOrderStyle(workbook, false, 8, HorizontalAlignment.CENTER, BorderStyle.NONE);
        smallCenterShrink.setShrinkToFit(true);
        CellStyle tableHeader = createCashOrderStyle(workbook, false, 10, HorizontalAlignment.CENTER, BorderStyle.THIN);
        CellStyle tableBoldCenter = createCashOrderStyle(workbook, true, 10, HorizontalAlignment.CENTER, BorderStyle.THIN);
        CellStyle tableBoldRight = createCashOrderStyle(workbook, true, 10, HorizontalAlignment.RIGHT, BorderStyle.THIN);
        CellStyle tableLeft = createCashOrderStyle(workbook, false, 10, HorizontalAlignment.LEFT, BorderStyle.THIN);
        CellStyle receiptTitle = createCashOrderStyle(workbook, true, 13, HorizontalAlignment.CENTER, BorderStyle.NONE);

        String orderShortName = firstNonBlank(header.get("order_short_name"), "ПКО");
        String orderTitle = "ПКО".equalsIgnoreCase(orderShortName)
                ? "ПРИХОДНЫЙ КАССОВЫЙ ОРДЕР"
                : "РАСХОДНЫЙ КАССОВЫЙ ОРДЕР";
        String receiptTitleText = "ПКО".equalsIgnoreCase(orderShortName) ? "КВИТАНЦИЯ" : "КВИТАНЦИЯ";
        String organization = firstNonBlank(header.get("organization_name"), "ОсОО \"СК Дрим Хаус\"");
        String orderNumber = firstNonBlank(header.get("order_number"), header.get("id"));
        String orderDate = stringValue(header.get("order_date"));
        String projectName = stringValue(header.get("project_name"));
        String blockName = stringValue(header.get("block_name"));
        String articleName = stringValue(header.get("article_name"));
        String titleText = stringValue(header.get("title"));
        String counterparty = stringValue(header.get("counterparty_name"));
        String counterpartyInn = stringValue(header.get("counterparty_inn"));
        String basis = firstNonBlank(header.get("basis"), titleText);
        String currency = firstNonBlank(header.get("currency"), "KGS");
        double amount = toDouble(header.get("amount"));
        double amountKgs = toDouble(firstNonNull(header.get("amount_kgs"), amount));
        String amountLabel = formatSmartNumber(amount) + " " + currency;
        String amountKgsLabel = formatSmartNumber(amountKgs);
        String amountWords = cashOrderAmountToWords(amount, currency);
        String receiptIntro = "ПКО".equalsIgnoreCase(orderShortName) ? "Принято от: " : "Выдано: ";
        String context = (projectName + (blockName.isBlank() ? "" : " / " + blockName)).trim();

        writeMerged(sheet, 1, 16, 23, "Унифицированная форма КО-1", center);
        writeMerged(sheet, 2, 16, 23, "Утверждена постановлением Нацстаткомитета", smallCenterShrink);
        writeMerged(sheet, 3, 16, 23, "Кыргызской Республики от 07.04.03\n№ 4", smallCenterShrink);

        writeMerged(sheet, 5, 2, 15, organization, title);
        writeMerged(sheet, 7, 7, 15, "организация", smallCenter);
        writeMerged(sheet, 9, 7, 15, "подразделение", smallCenter);

        writeMerged(sheet, 5, 20, 23, "Код", tableHeader);
        writeMerged(sheet, 6, 16, 19, "Форма по ГКУД", center);
        writeMerged(sheet, 6, 20, 23, "0310001", tableBoldCenter);
        writeMerged(sheet, 7, 16, 19, "по ОКПО", center);
        writeMerged(sheet, 7, 20, 23, "31431840", tableBoldCenter);
        writeMerged(sheet, 11, 2, 13, orderTitle, orderTitleStyle);
        writeMerged(sheet, 10, 14, 19, "Номер документа", tableHeader);
        writeMerged(sheet, 10, 20, 23, "Дата составления", tableHeader);
        writeMerged(sheet, 11, 14, 19, orderShortName + "-" + orderNumber, tableBoldCenter);
        writeMerged(sheet, 11, 20, 23, orderDate, tableBoldCenter);

        writeMerged(sheet, 13, 2, 5, "Дебет", tableHeader);
        writeMerged(sheet, 13, 6, 15, "Кредит", tableHeader);
        writeMerged(sheet, 13, 16, 20, "Сумма,", tableHeader);
        writeMerged(sheet, 13, 21, 23, "Код", tableHeader);
        writeMerged(sheet, 14, 6, 9, "код\nструктурного\nподразделения", tableHeader);
        writeMerged(sheet, 14, 10, 13, "корр.счет,\nсубсчет", tableHeader);
        writeMerged(sheet, 14, 14, 15, "код ан.\nучета", tableHeader);
        writeMerged(sheet, 14, 16, 20, currency, tableHeader);
        writeMerged(sheet, 14, 21, 23, "целевого\nназначения", tableHeader);
        writeMerged(sheet, 15, 2, 5, "1110", tableLeft);
        writeMerged(sheet, 15, 6, 9, "", tableLeft);
        writeMerged(sheet, 15, 10, 13, "3210", tableLeft);
        writeMerged(sheet, 15, 14, 15, "", tableLeft);
        writeMerged(sheet, 15, 16, 20, amountLabel, tableBoldRight);
        writeMerged(sheet, 15, 21, 23, "", tableLeft);

        writeMerged(sheet, 17, 2, 5, "Принято от:", normal);
        writeMerged(sheet, 17, 6, 23, counterparty, normal);
        writeMerged(sheet, 18, 6, 23, "фамилия, имя, отчество", smallCenter);
        writeMerged(sheet, 19, 2, 5, "Основание:", normal);
        writeMerged(sheet, 19, 6, 23, basis, normal);
        writeMerged(sheet, 21, 2, 5, "Сумма:", normal);
        writeMerged(sheet, 21, 6, 23, amountWords, normal);
        writeMerged(sheet, 22, 14, 23, "прописью", smallCenter);
        writeMerged(sheet, 23, 2, 5, "В том числе:", normal);
        writeMerged(sheet, 25, 2, 5, "Приложение:", normal);

        writeMerged(sheet, 27, 2, 7, "Главный бухгалтер", bold);
        writeMerged(sheet, 28, 8, 12, "подпись", smallCenter);
        writeMerged(sheet, 28, 14, 23, "расшифровка подписи", smallCenter);
        writeMerged(sheet, 30, 2, 7, "Получил кассир", bold);
        writeMerged(sheet, 31, 8, 12, "подпись", smallCenter);
        writeMerged(sheet, 31, 14, 23, "расшифровка подписи", smallCenter);
        addBottomBorder(sheet, 28, 8, 12);
        addBottomBorder(sheet, 28, 14, 23);
        addBottomBorder(sheet, 31, 8, 12);
        addBottomBorder(sheet, 31, 14, 23);

        for (int row = 1; row <= 32; row++) {
            addRightDashedBorder(sheet, row, 24);
        }

        writeMerged(sheet, 2, 25, 33, organization, receiptTitle);
        writeMerged(sheet, 3, 29, 33, "организация", smallCenter);
        writeMerged(sheet, 5, 27, 33, receiptTitleText, title);
        writeMerged(sheet, 6, 27, 33, "к " + orderShortName + " № " + orderShortName + "-" + orderNumber, center);
        writeMerged(sheet, 7, 27, 33, "от " + orderDate, tableBoldCenter);
        writeMerged(sheet, 10, 25, 33, receiptIntro + counterparty, normal);
        writeMerged(sheet, 13, 25, 33, "Основание:", normal);
        writeMerged(sheet, 14, 25, 33, basis, normal);
        writeMerged(sheet, 16, 25, 27, "Сумма:", normal);
        writeMerged(sheet, 16, 28, 33, amountWords, bold);
        writeMerged(sheet, 17, 25, 27, currency, normal);
        writeMerged(sheet, 17, 28, 33, "цифрами: " + amountLabel, smallCenter);
        writeMerged(sheet, 19, 25, 33, amountWords, normal);
        writeMerged(sheet, 23, 25, 28, "М.П. (штамп)", bold);
        writeMerged(sheet, 23, 29, 33, orderDate, bold);
        writeMerged(sheet, 27, 25, 30, "Главный бухгалтер", bold);
        writeMerged(sheet, 29, 25, 28, "подпись", smallCenter);
        writeMerged(sheet, 29, 29, 33, "расшифровка подписи", smallCenter);
        writeMerged(sheet, 30, 25, 30, "Кассир", bold);
        writeMerged(sheet, 32, 25, 28, "подпись", smallCenter);
        writeMerged(sheet, 32, 29, 33, "расшифровка подписи", smallCenter);
        addBottomBorder(sheet, 29, 25, 28);
        addBottomBorder(sheet, 29, 29, 33);
        addBottomBorder(sheet, 32, 25, 28);
        addBottomBorder(sheet, 32, 29, 33);

        writeMerged(sheet, 34, 2, 23, "Объект: " + context + " | Статья: " + articleName + (counterpartyInn.isBlank() ? "" : " | ИНН: " + counterpartyInn), normal);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        return out.toByteArray();
    }

    public byte[] generateForm29(ReportData reportData, String format) throws Exception {
        if (!"html".equalsIgnoreCase(format) && !"xlsx".equalsIgnoreCase(format) && (format == null || !"__legacy__".equalsIgnoreCase(format))) {
            InputStream jrxmlStream =
                    new ClassPathResource("reports/Form29Report.jrxml").getInputStream();

            JasperReport report = JasperCompileManager.compileReport(jrxmlStream);

            Map data = reportData.data();
            Map header = reportData.header();
            List<Map<String, Object>> materials = (List<Map<String, Object>>) data.getOrDefault("materials", List.of());
            List<Map<String, Object>> sections = (List<Map<String, Object>>) data.getOrDefault("sections", List.of());
            List<Map<String, Object>> totals = (List<Map<String, Object>>) data.getOrDefault("totals", List.of());

            List<Map<String, Object>> bodyRows = new ArrayList<>();
            List<Map<String, Object>> totalRows = new ArrayList<>();

            int materialOrder = 0;
            Map<Integer, Integer> materialOrderMap = new LinkedHashMap<>();
            for (Map<String, Object> material : materials) {
                materialOrderMap.put(toInt(material.get("material_id")), materialOrder++);
            }

            int sectionOrder = 0;
            for (Map<String, Object> section : sections) {
                String sectionTitle = buildForm29SectionTitle(section);
                String sectionBucket = form29Bucket(sectionOrder, 0, "section");

                for (Map<String, Object> material : materials) {
                    int currentMaterialOrder = materialOrderMap.getOrDefault(toInt(material.get("material_id")), 0);
                    bodyRows.add(buildForm29BodyRow(sectionBucket, sectionTitle, "section", material, currentMaterialOrder, "01|norm", "\u043d\u043e\u0440\u043c", ""));
                    bodyRows.add(buildForm29BodyRow(sectionBucket, sectionTitle, "section", material, currentMaterialOrder, "02|fact", "\u0444\u0430\u043a\u0442", ""));
                }

                List<Map<String, Object>> rows = (List<Map<String, Object>>) section.getOrDefault("rows", List.of());
                int rowOrder = 1;
                for (Map<String, Object> row : rows) {
                    Map<Integer, Map<String, Object>> cellByMaterial = new LinkedHashMap<>();
                    List<Map<String, Object>> cells = (List<Map<String, Object>>) row.getOrDefault("cells", List.of());
                    for (Map<String, Object> cell : cells) {
                        cellByMaterial.put(toInt(cell.get("material_id")), cell);
                    }

                    String rowBucket = form29Bucket(sectionOrder, rowOrder, "work");
                    String workLabel = stringValue(row.get("work_name"));

                    for (Map<String, Object> material : materials) {
                        int materialId = toInt(material.get("material_id"));
                        int currentMaterialOrder = materialOrderMap.getOrDefault(materialId, 0);
                        Map<String, Object> cell = cellByMaterial.get(materialId);

                        bodyRows.add(buildForm29BodyRow(
                                rowBucket,
                                workLabel,
                                "work",
                                material,
                                currentMaterialOrder,
                                "01|norm",
                                "\u043d\u043e\u0440\u043c",
                                formatForm29Norm(cell)
                        ));
                        bodyRows.add(buildForm29BodyRow(
                                rowBucket,
                                workLabel,
                                "work",
                                material,
                                currentMaterialOrder,
                                "02|fact",
                                "\u0444\u0430\u043a\u0442",
                                formatSmartNumber(cell == null ? 0 : toDouble(cell.get("actual_quantity")))
                        ));
                    }

                    rowOrder++;
                }

                sectionOrder++;
            }

            List<Map<String, Object>> totalDefinitions = List.of(
                    Map.of("label", "\u0418\u0422\u041e\u0413\u041e \u0420\u0410\u0421\u0425\u041e\u0414 \u041f\u041e \u041d\u041e\u0420\u041c\u0415", "field", "planned_quantity"),
                    Map.of("label", "\u0424\u0410\u041a\u0422\u0418\u0427\u0415\u0421\u041a\u0418\u0419 \u0420\u0410\u0421\u0425\u041e\u0414", "field", "actual_quantity"),
                    Map.of("label", "\u042d\u041a\u041e\u041d\u041e\u041c\u0418\u042f (+), \u041f\u0415\u0420\u0415\u0420\u0410\u0421\u0425\u041e\u0414 (-)", "field", "deviation_quantity"),
                    Map.of("label", "\u0421\u041f\u0418\u0421\u0410\u0422\u042c \u041d\u0410 \u0421\u0415\u0411\u0415\u0421\u0422\u041e\u0418\u041c\u041e\u0421\u0422\u042c", "field", "actual_quantity")
            );

            Map<Integer, Map<String, Object>> totalsByMaterial = new LinkedHashMap<>();
            for (Map<String, Object> total : totals) {
                totalsByMaterial.put(toInt(total.get("material_id")), total);
            }

            int totalOrder = 0;
            for (Map<String, Object> totalDefinition : totalDefinitions) {
                String totalBucket = String.format("%03d|%s", totalOrder, totalDefinition.get("label"));
                String totalField = stringValue(totalDefinition.get("field"));

                for (Map<String, Object> material : materials) {
                    int materialId = toInt(material.get("material_id"));
                    int currentMaterialOrder = materialOrderMap.getOrDefault(materialId, 0);
                    Map<String, Object> total = totalsByMaterial.get(materialId);
                    double totalValue = total == null ? 0 : toDouble(total.get(totalField));
                    String factValue = "deviation_quantity".equals(totalField)
                            ? formatSignedSmartNumber(totalValue)
                            : formatSmartNumber(totalValue);

                    totalRows.add(buildForm29TotalRow(totalBucket, stringValue(totalDefinition.get("label")), material, currentMaterialOrder, "01|norm", "\u043d\u043e\u0440\u043c", "X"));
                    totalRows.add(buildForm29TotalRow(totalBucket, stringValue(totalDefinition.get("label")), material, currentMaterialOrder, "02|fact", "\u0444\u0430\u043a\u0442", factValue));
                }

                totalOrder++;
            }

            Map<String, Object> params = new HashMap<>();
            params.put("customerName", "\u041e\u0441\u041e\u041e \"\u0421\u041a \u0414\u0440\u0438\u043c \u0425\u0430\u0443\u0441\"");
            params.put("projectName", stringValue(header.get("project_name")));
            params.put("blockName", stringValue(header.get("block_name")));
            params.put("periodLabel", stringValue(header.get("period_label")));
            params.put("bodyDataSource", new JRMapCollectionDataSource(new ArrayList<>(bodyRows)));
            params.put("totalsDataSource", new JRMapCollectionDataSource(new ArrayList<>(totalRows)));

            JasperPrint print = JasperFillManager.fillReport(report, params, new JREmptyDataSource(1));

            return switch (normalizeFormat(format, "pdf")) {
                case "xlsx" -> exportXlsx(print);
                case "docx" -> exportDocx(print);
                case "html" -> exportHtml(print);
                default -> JasperExportManager.exportReportToPdf(print);
            };
        }

        if ("html".equalsIgnoreCase(format)) {
            return generateForm29Html(reportData);
        }
        if (!"xlsx".equalsIgnoreCase(format)) {
            throw new IllegalArgumentException("Unsupported Form29 format: " + format);
        }

        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> materials = (List<Map<String, Object>>) data.getOrDefault("materials", List.of());
        List<Map<String, Object>> sections = (List<Map<String, Object>>) data.getOrDefault("sections", List.of());
        List<Map<String, Object>> totals = (List<Map<String, Object>>) data.getOrDefault("totals", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("\u0424\u043e\u0440\u043c\u0430 29");

            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            CellStyle labelStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, null);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle sectionStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle integerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            integerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            CellStyle totalIntegerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            totalIntegerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            CellStyle totalDecimalStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            totalDecimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            CellStyle additionalHeaderStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(additionalHeaderStyle, "F5E7D3");
            CellStyle additionalCenterStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(additionalCenterStyle, "F5E7D3");
            CellStyle additionalIntegerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(additionalIntegerStyle, "F5E7D3");
            additionalIntegerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            CellStyle additionalDecimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(additionalDecimalStyle, "F5E7D3");
            additionalDecimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            CellStyle additionalTotalIntegerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(additionalTotalIntegerStyle, "F5E7D3");
            additionalTotalIntegerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            CellStyle additionalTotalDecimalStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(additionalTotalDecimalStyle, "F5E7D3");
            additionalTotalDecimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            CellStyle estimateHeaderStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(estimateHeaderStyle, "E4EEDF");
            CellStyle estimateCenterStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(estimateCenterStyle, "E4EEDF");
            CellStyle estimateIntegerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(estimateIntegerStyle, "E4EEDF");
            estimateIntegerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            CellStyle estimateDecimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(estimateDecimalStyle, "E4EEDF");
            estimateDecimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            CellStyle estimateTotalIntegerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(estimateTotalIntegerStyle, "E4EEDF");
            estimateTotalIntegerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            CellStyle estimateTotalDecimalStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(estimateTotalDecimalStyle, "E4EEDF");
            estimateTotalDecimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));

            int lastColumn = Math.max(2, materials.size() * 2 + 2);
            int rowIndex = 0;

            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u0417\u0430\u043a\u0430\u0437\u0447\u0438\u043a: \u041e\u0441\u041e\u041e \"\u0421\u041a \u0414\u0440\u0438\u043c \u0425\u0430\u0443\u0441\"", labelStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u041e\u0431\u044a\u0435\u043a\u0442: " + stringValue(header.get("project_name")) + ", \u0431\u043b\u043e\u043a " + stringValue(header.get("block_name")), labelStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u041e\u0422\u0427\u0415\u0422 \u0424-29", titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u0417\u0430 " + stringValue(header.get("period_label")), titleStyle);
            rowIndex++;

            Row topHeader = sheet.createRow(rowIndex++);
            topHeader.setHeightInPoints(42);
            writeCell(topHeader, 0, "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0442", headerStyle);

            writeCell(topHeader, 1, "\u0415\u0434.\n\u0438\u0437\u043c.", headerStyle);
            writeCell(topHeader, 2, "\u041e\u0431\u044a\u0435\u043c", headerStyle);


            int materialColumn = 3;
            for (Map<String, Object> material : materials) {
                CellStyle materialHeaderStyle = resolveMaterialHeaderStyle(
                        material,
                        headerStyle,
                        additionalHeaderStyle,
                        estimateHeaderStyle
                );
                mergeCellsIfNeeded(sheet, topHeader.getRowNum(), materialColumn, materialColumn + 1);
                writeCell(topHeader, materialColumn, stringValue(material.get("material_name")), materialHeaderStyle);
                writeCell(topHeader, materialColumn + 1, "", materialHeaderStyle);
                materialColumn += 2;
            }

            Row secondHeader = sheet.createRow(rowIndex++);
            writeCell(secondHeader, 0, "", headerStyle);
            writeCell(secondHeader, 1, "", headerStyle);
            writeCell(secondHeader, 2, "", headerStyle);

            materialColumn = 3;
            for (Map<String, Object> material : materials) {
                CellStyle materialHeaderStyle = resolveMaterialHeaderStyle(
                        material,
                        headerStyle,
                        additionalHeaderStyle,
                        estimateHeaderStyle
                );
                writeCell(secondHeader, materialColumn, "\u043d\u043e\u0440\u043c", materialHeaderStyle);
                writeCell(secondHeader, materialColumn + 1, "\u0444\u0430\u043a\u0442", materialHeaderStyle);
                materialColumn += 2;
            }

            Row unitsRow = sheet.createRow(rowIndex++);
            writeCell(unitsRow, 0, "", headerStyle);
            writeCell(unitsRow, 1, "", headerStyle);
            writeCell(unitsRow, 2, "", headerStyle);

            materialColumn = 3;
            for (Map<String, Object> material : materials) {
                CellStyle materialHeaderStyle = resolveMaterialHeaderStyle(
                        material,
                        headerStyle,
                        additionalHeaderStyle,
                        estimateHeaderStyle
                );
                String unitName = stringValue(material.get("unit_name"));
                writeCell(unitsRow, materialColumn, unitName, materialHeaderStyle);
                writeCell(unitsRow, materialColumn + 1, unitName, materialHeaderStyle);
                materialColumn += 2;
            }

            for (Map<String, Object> section : sections) {
                Row sectionRow = sheet.createRow(rowIndex++);
                for (int c = 0; c <= lastColumn; c++) {
                    writeCell(sectionRow, c, "", sectionStyle);
                }

                String sectionTitle = stringValue(section.get("stage_name"));
                String subsectionName = stringValue(section.get("subsection_name"));
                if (!subsectionName.isBlank()) {
                    sectionTitle = sectionTitle + " / " + subsectionName;
                }

                writeCell(sectionRow, 0, sectionTitle, sectionStyle);
                mergeCellsIfNeeded(sheet, sectionRow.getRowNum(), 0, 2);

                List<Map<String, Object>> rows = (List<Map<String, Object>>) section.getOrDefault("rows", List.of());
                for (Map<String, Object> row : rows) {
                    Row excelRow = sheet.createRow(rowIndex++);
                    writeCell(excelRow, 0, stringValue(row.get("work_name")), textStyle);
                    writeCell(excelRow, 1, stringValue(row.get("unit_name")), centerStyle);
                    writeSmartNumericCell(excelRow, 2, toDouble(row.get("quantity")), integerStyle, decimalStyle);

                    Map<Integer, Map<String, Object>> cellByMaterial = new LinkedHashMap<>();
                    List<Map<String, Object>> cells = (List<Map<String, Object>>) row.getOrDefault("cells", List.of());
                    for (Map<String, Object> cell : cells) {
                        cellByMaterial.put(toInt(cell.get("material_id")), cell);
                    }

                    materialColumn = 3;
                    for (Map<String, Object> material : materials) {
                        int materialId = toInt(material.get("material_id"));
                        Map<String, Object> cell = cellByMaterial.get(materialId);
                        CellStyle materialCenterStyle = resolveMaterialCenterStyle(
                                material,
                                centerStyle,
                                additionalCenterStyle,
                                estimateCenterStyle
                        );
                        CellStyle materialIntegerStyle = resolveMaterialIntegerStyle(
                                material,
                                integerStyle,
                                additionalIntegerStyle,
                                estimateIntegerStyle
                        );
                        CellStyle materialDecimalStyle = resolveMaterialDecimalStyle(
                                material,
                                decimalStyle,
                                additionalDecimalStyle,
                                estimateDecimalStyle
                        );

                        if (cell == null) {
                            writeCell(excelRow, materialColumn, "", materialCenterStyle);
                            writeSmartNumericCell(excelRow, materialColumn + 1, 0, materialIntegerStyle, materialDecimalStyle);
                        } else {
                            Object normValue = cell.get("norm_value");
                            if (normValue instanceof Number) {
                                writeSmartNumericCell(excelRow, materialColumn, toDouble(normValue), materialIntegerStyle, materialDecimalStyle);
                            } else {
                                writeCell(excelRow, materialColumn, stringValue(normValue), materialCenterStyle);
                            }
                            writeSmartNumericCell(excelRow, materialColumn + 1, toDouble(cell.get("actual_quantity")), materialIntegerStyle, materialDecimalStyle);
                        }

                        materialColumn += 2;
                    }
                }
            }

            Map<Integer, Map<String, Object>> totalsByMaterial = new LinkedHashMap<>();
            for (Map<String, Object> total : totals) {
                totalsByMaterial.put(toInt(total.get("material_id")), total);
            }
            rowIndex = writeTotalsRow(sheet, rowIndex, materials, totalsByMaterial, "\u0418\u0422\u041e\u0413\u041e \u0420\u0410\u0421\u0425\u041e\u0414 \u041f\u041e \u041d\u041e\u0420\u041c\u0415", "planned_quantity", totalIntegerStyle, totalDecimalStyle, additionalTotalIntegerStyle, additionalTotalDecimalStyle, estimateTotalIntegerStyle, estimateTotalDecimalStyle, centerStyle, additionalCenterStyle, estimateCenterStyle);
            rowIndex = writeTotalsRow(sheet, rowIndex, materials, totalsByMaterial, "\u0424\u0410\u041a\u0422\u0418\u0427\u0415\u0421\u041a\u0418\u0419 \u0420\u0410\u0421\u0425\u041e\u0414", "actual_quantity", totalIntegerStyle, totalDecimalStyle, additionalTotalIntegerStyle, additionalTotalDecimalStyle, estimateTotalIntegerStyle, estimateTotalDecimalStyle, centerStyle, additionalCenterStyle, estimateCenterStyle);
            rowIndex = writeTotalsRow(sheet, rowIndex, materials, totalsByMaterial, "\u042d\u041a\u041e\u041d\u041e\u041c\u0418\u042f (+), \u041f\u0415\u0420\u0415\u0420\u0410\u0421\u0425\u041e\u0414 (-)", "deviation_quantity", totalIntegerStyle, totalDecimalStyle, additionalTotalIntegerStyle, additionalTotalDecimalStyle, estimateTotalIntegerStyle, estimateTotalDecimalStyle, centerStyle, additionalCenterStyle, estimateCenterStyle);
            rowIndex = writeTotalsRow(sheet, rowIndex, materials, totalsByMaterial, "\u0421\u041f\u0418\u0421\u0410\u0422\u042c \u041d\u0410 \u0421\u0415\u0411\u0415\u0421\u0422\u041e\u0418\u041c\u041e\u0421\u0422\u042c", "actual_quantity", totalIntegerStyle, totalDecimalStyle, additionalTotalIntegerStyle, additionalTotalDecimalStyle, estimateTotalIntegerStyle, estimateTotalDecimalStyle, centerStyle, additionalCenterStyle, estimateCenterStyle);

            rowIndex++;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u041d\u0430\u0447.\u041f\u0422\u041e:", labelStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u0422\u0435\u0445.\u0434\u0438\u0440\u0435\u043a\u0442\u043e\u0440:", labelStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u041f\u0440\u043e\u0440\u0430\u0431:", labelStyle);

            sheet.setColumnWidth(0, 38 * 256);
            sheet.setColumnWidth(1, 5 * 256);
            sheet.setColumnWidth(2, 7 * 256);
            for (int c = 3; c <= lastColumn; c++) {
                sheet.setColumnWidth(c, 6 * 256);
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }

    private byte[] generateForm29Html(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> materials = (List<Map<String, Object>>) data.getOrDefault("materials", List.of());
        List<Map<String, Object>> sections = (List<Map<String, Object>>) data.getOrDefault("sections", List.of());
        List<Map<String, Object>> totals = (List<Map<String, Object>>) data.getOrDefault("totals", List.of());

        Map<Integer, Map<String, Object>> totalsByMaterial = new LinkedHashMap<>();
        for (Map<String, Object> total : totals) {
            totalsByMaterial.put(toInt(total.get("material_id")), total);
        }

        StringBuilder html = new StringBuilder();
        html.append("<!doctype html><html><head><meta charset='UTF-8'>");
        html.append("<style>");
        html.append("body{font-family:'Times New Roman',serif;margin:16px;color:#111;} ");
        html.append(".meta{font-size:16px;margin-bottom:4px;} ");
        html.append(".title{font-weight:700;text-align:center;margin:8px 0 0;font-size:20px;} ");
        html.append(".period{font-weight:700;text-align:center;margin:2px 0 12px;font-size:18px;text-decoration:underline;} ");
        html.append("table{border-collapse:collapse;width:max-content;min-width:100%;} ");
        html.append("th,td{border:1px solid #111;padding:3px 4px;font-size:12px;vertical-align:middle;} ");
        html.append("th{font-weight:700;text-align:center;} ");
        html.append(".work-col{min-width:320px;text-align:left;} ");
        html.append(".unit-col{min-width:44px;text-align:center;} ");
        html.append(".qty-col{min-width:72px;text-align:center;} ");
        html.append(".section{font-weight:700;text-align:center;} ");
        html.append(".base{background:#ffffff;} ");
        html.append(".additional{background:#f5e7d3;} ");
        html.append(".estimate{background:#e4eedf;} ");
        html.append(".label{font-weight:700;text-align:center;} ");
        html.append(".num{text-align:center;min-width:52px;} ");
        html.append(".footer{margin-top:2px;font-weight:700;font-style:italic;} ");
        html.append("</style></head><body>");
        html.append("<div class='meta'>\u0417\u0430\u043a\u0430\u0437\u0447\u0438\u043a: \u041e\u0441\u041e\u041e &quot;\u0421\u041a \u0414\u0440\u0438\u043c \u0425\u0430\u0443\u0441&quot;</div>");
        html.append("<div class='meta'>\u041e\u0431\u044a\u0435\u043a\u0442: ").append(escapeHtml(stringValue(header.get("project_name"))))
                .append(", \u0431\u043b\u043e\u043a ").append(escapeHtml(stringValue(header.get("block_name")))).append("</div>");
        html.append("<div class='title'>\u041e\u0422\u0427\u0415\u0422 \u0424-29</div>");
        html.append("<div class='period'>\u0417\u0430 ").append(escapeHtml(stringValue(header.get("period_label")))).append("</div>");

        html.append("<table>");
        html.append("<thead>");
        html.append("<tr>");
        html.append("<th class='base work-col' rowspan='3'>\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0442</th>");
        html.append("<th class='base unit-col' rowspan='3'>\u0415\u0434.<br>\u0438\u0437\u043c.</th>");
        html.append("<th class='base qty-col' rowspan='3'>\u041e\u0431\u044a\u0435\u043c</th>");
        for (Map<String, Object> material : materials) {
            String materialClass = materialCssClass(material);
            html.append("<th class='").append(materialClass).append("' colspan='2'>")
                    .append(escapeHtml(stringValue(material.get("material_name"))))
                    .append("</th>");
        }
        html.append("</tr><tr>");
        for (Map<String, Object> material : materials) {
            String materialClass = materialCssClass(material);
            html.append("<th class='").append(materialClass).append("'>\u043d\u043e\u0440\u043c</th>");
            html.append("<th class='").append(materialClass).append("'>\u0444\u0430\u043a\u0442</th>");
        }
        html.append("</tr><tr>");
        for (Map<String, Object> material : materials) {
            String materialClass = materialCssClass(material);
            String unitName = escapeHtml(stringValue(material.get("unit_name")));
            html.append("<th class='").append(materialClass).append("'>").append(unitName).append("</th>");
            html.append("<th class='").append(materialClass).append("'>").append(unitName).append("</th>");
        }
        html.append("</tr></thead><tbody>");

        for (Map<String, Object> section : sections) {
            String sectionTitle = stringValue(section.get("stage_name"));
            String subsectionName = stringValue(section.get("subsection_name"));
            if (!subsectionName.isBlank()) {
                sectionTitle = sectionTitle + " / " + subsectionName;
            }
            html.append("<tr><td class='section' colspan='").append(Math.max(3, materials.size() * 2 + 3)).append("'>")
                    .append(escapeHtml(sectionTitle))
                    .append("</td></tr>");

            List<Map<String, Object>> rows = (List<Map<String, Object>>) section.getOrDefault("rows", List.of());
            for (Map<String, Object> row : rows) {
                html.append("<tr>");
                html.append("<td class='work-col'>").append(escapeHtml(stringValue(row.get("work_name")))).append("</td>");
                html.append("<td class='unit-col'>").append(escapeHtml(stringValue(row.get("unit_name")))).append("</td>");
                html.append("<td class='qty-col'>").append(escapeHtml(formatSmartNumber(toDouble(row.get("quantity"))))).append("</td>");

                Map<Integer, Map<String, Object>> cellByMaterial = new LinkedHashMap<>();
                List<Map<String, Object>> cells = (List<Map<String, Object>>) row.getOrDefault("cells", List.of());
                for (Map<String, Object> cell : cells) {
                    cellByMaterial.put(toInt(cell.get("material_id")), cell);
                }

                for (Map<String, Object> material : materials) {
                    String materialClass = materialCssClass(material);
                    Map<String, Object> cell = cellByMaterial.get(toInt(material.get("material_id")));
                    if (cell == null) {
                        html.append("<td class='num ").append(materialClass).append("'></td>");
                        html.append("<td class='num ").append(materialClass).append("'>0</td>");
                        continue;
                    }

                    Object normValue = cell.get("norm_value");
                    String normText = normValue instanceof Number
                            ? formatSmartNumber(toDouble(normValue))
                            : escapeHtml(stringValue(normValue));

                    html.append("<td class='num ").append(materialClass).append("'>").append(normText).append("</td>");
                    html.append("<td class='num ").append(materialClass).append("'>")
                            .append(formatSmartNumber(toDouble(cell.get("actual_quantity"))))
                            .append("</td>");
                }

                html.append("</tr>");
            }
        }

        appendHtmlTotalsRow(html, materials, totalsByMaterial, "\u0418\u0422\u041e\u0413\u041e \u0420\u0410\u0421\u0425\u041e\u0414 \u041f\u041e \u041d\u041e\u0420\u041c\u0415", "planned_quantity");
        appendHtmlTotalsRow(html, materials, totalsByMaterial, "\u0424\u0410\u041a\u0422\u0418\u0427\u0415\u0421\u041a\u0418\u0419 \u0420\u0410\u0421\u0425\u041e\u0414", "actual_quantity");
        appendHtmlTotalsRow(html, materials, totalsByMaterial, "\u042d\u041a\u041e\u041d\u041e\u041c\u0418\u042f (+), \u041f\u0415\u0420\u0415\u0420\u0410\u0421\u0425\u041e\u0414 (-)", "deviation_quantity");
        appendHtmlTotalsRow(html, materials, totalsByMaterial, "\u0421\u041f\u0418\u0421\u0410\u0422\u042c \u041d\u0410 \u0421\u0415\u0411\u0415\u0421\u0422\u041e\u0418\u041c\u041e\u0421\u0422\u042c", "actual_quantity");
        html.append("</tbody></table>");

        html.append("<div class='footer'>\u041d\u0430\u0447.\u041f\u0422\u041e:</div>");
        html.append("<div class='footer'>\u0422\u0435\u0445.\u0434\u0438\u0440\u0435\u043a\u0442\u043e\u0440:</div>");
        html.append("<div class='footer'>\u041f\u0440\u043e\u0440\u0430\u0431:</div>");
        html.append("</body></html>");
        return html.toString().getBytes(StandardCharsets.UTF_8);
    }


    public byte[] generateForm2(ReportData reportData, String format) throws Exception {
        if (!"html".equalsIgnoreCase(format) && !"xlsx".equalsIgnoreCase(format)) {
            InputStream jrxmlStream =
                    new ClassPathResource("reports/Form2Report.jrxml").getInputStream();

            JasperReport report = JasperCompileManager.compileReport(jrxmlStream);

            Map data = reportData.data();
            Map header = reportData.header();
            List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

            List<Map<String, Object>> printableRows = buildForm2PrintableRows(rows);
            JRMapCollectionDataSource dataSource = new JRMapCollectionDataSource(new ArrayList<>(printableRows));

            Map<String, Object> params = new HashMap<>();
            params.put("customerName", stringValue(header.get("customer_name")));
            params.put("projectName", stringValue(header.get("project_name")));
            params.put("projectAddress", stringValue(header.get("project_address")));
            params.put("constructionType", stringValue(header.get("construction_type")));
            params.put("reportNumber", stringValue(header.get("report_number")));
            params.put("reportDateLabel", stringValue(header.get("report_date_label")));
            params.put("periodFromLabel", stringValue(header.get("period_from_label")));
            params.put("periodToLabel", stringValue(header.get("period_to_label")));
            params.put("periodLabel", stringValue(header.get("period_label")));
            params.put("foremanName", stringValue(header.get("foreman_name")));

            JasperPrint print = JasperFillManager.fillReport(report, params, dataSource);

            return switch (normalizeFormat(format, "pdf")) {
                case "xlsx" -> exportXlsx(print);
                case "docx" -> exportDocx(print);
                case "html" -> exportHtml(print);
                default -> JasperExportManager.exportReportToPdf(print);
            };
        }

        if ("html".equalsIgnoreCase(format)) {
            return generateForm2Html(reportData);
        }
        if (!"xlsx".equalsIgnoreCase(format)) {
            throw new IllegalArgumentException("Unsupported Form2 format: " + format);
        }

        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("\u0424\u043e\u0440\u043c\u0430 2");

            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            CellStyle boldLeftStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, null);
            CellStyle italicLeftStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            Font italicFont = workbook.createFont();
            italicFont.setFontName("Times New Roman");
            italicFont.setItalic(true);
            italicFont.setFontHeightInPoints((short) 11);
            italicLeftStyle.setFont(italicFont);

            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle leftStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle integerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            integerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            CellStyle percentStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            percentStyle.setDataFormat(workbook.createDataFormat().getFormat("0.##"));
            CellStyle materialNameStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            Font materialFont = workbook.createFont();
            materialFont.setFontName("Arial");
            materialFont.setItalic(true);
            materialNameStyle.setFont(materialFont);
            CellStyle workNameStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, BorderStyle.THIN);

            int lastColumn = 13;
            int rowIndex = 0;

            Row row0 = sheet.createRow(rowIndex++);
            writeCell(row0, 0, "\u0417\u0430\u043a\u0430\u0437\u0447\u0438\u043a:", boldLeftStyle);
            writeCell(row0, 1, stringValue(header.get("customer_name")), italicLeftStyle);
            mergeCellsIfNeeded(sheet, row0.getRowNum(), 1, 5);
            writeCell(row0, 7, "№", headerStyle);
            mergeCellsIfNeeded(sheet, row0.getRowNum(), 7, 7);
            writeCell(row0, 8, "\u0414\u0430\u0442\u0430", headerStyle);
            mergeCellsIfNeeded(sheet, row0.getRowNum(), 8, 9);
            writeCell(row0, 10, "\u041e\u0442\u0447\u0435\u0442\u043d\u044b\u0439 \u043f\u0435\u0440\u0438\u043e\u0434:", headerStyle);
            mergeCellsIfNeeded(sheet, row0.getRowNum(), 10, 13);

            Row row1 = sheet.createRow(rowIndex++);
            writeCell(row1, 0, "\u041e\u0431\u044a\u0435\u043a\u0442:", boldLeftStyle);
            writeCell(row1, 1, stringValue(header.get("project_name")) + ", " + stringValue(header.get("project_address")), italicLeftStyle);
            mergeCellsIfNeeded(sheet, row1.getRowNum(), 1, 5);
            writeCell(row1, 7, "\u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0430", headerStyle);
            writeCell(row1, 8, "\u0441\u043e\u0441\u0442\u0430\u0432\u043b\u0435\u043d\u0438\u044f", headerStyle);
            mergeCellsIfNeeded(sheet, row1.getRowNum(), 8, 9);
            writeCell(row1, 10, "\u0441", headerStyle);
            writeCell(row1, 12, "\u043f\u043e", headerStyle);
            mergeCellsIfNeeded(sheet, row1.getRowNum(), 10, 11);
            mergeCellsIfNeeded(sheet, row1.getRowNum(), 12, 13);

            Row row2 = sheet.createRow(rowIndex++);
            writeCell(row2, 0, "\u0412\u0438\u0434 \u0421\u041c\u0420:", boldLeftStyle);
            writeCell(row2, 1, stringValue(header.get("construction_type")), italicLeftStyle);
            mergeCellsIfNeeded(sheet, row2.getRowNum(), 1, 5);
            writeCell(row2, 7, stringValue(header.get("report_number")), centerStyle);
            writeCell(row2, 8, stringValue(header.get("report_date_label")), centerStyle);
            mergeCellsIfNeeded(sheet, row2.getRowNum(), 8, 9);
            writeCell(row2, 10, stringValue(header.get("period_from_label")), centerStyle);
            mergeCellsIfNeeded(sheet, row2.getRowNum(), 10, 11);
            writeCell(row2, 12, stringValue(header.get("period_to_label")), centerStyle);
            mergeCellsIfNeeded(sheet, row2.getRowNum(), 12, 13);

            rowIndex++;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u0410\u041a\u0422", titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "\u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u044b\u0445 \u0440\u0430\u0431\u043e\u0442 \u0437\u0430 " + stringValue(header.get("period_label")), titleStyle);

            Row topHeader = sheet.createRow(rowIndex++);
            topHeader.setHeightInPoints(34);
            writeCell(topHeader, 0, "\u041d\u043e\u043c\u0435\u0440", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 0, 1);
            writeCell(topHeader, 2, "\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0442", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 2, 2);
            writeCell(topHeader, 3, "\u0435\u0434.\n\u0438\u0437\u043c.", headerStyle);
            writeCell(topHeader, 4, "\u043f\u043e \u0434\u043e\u0433\u043e\u0432\u043e\u0440\u0443", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 4, 6);
            writeCell(topHeader, 7, "\u0424\u0430\u043a\u0442\u0438\u0447\u0435\u0441\u043a\u043e\u0435 \u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u0435 \u0444\u0438\u0437\u0438\u0447\u0435\u0441\u043a\u0438\u0445 \u043e\u0431\u044a\u0435\u043c\u043e\u0432 \u0440\u0430\u0431\u043e\u0442 \u0438 \u0441\u0442\u043e\u0438\u043c\u043e\u0441\u0442\u0438:", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 7, 13);
            Row middleHeader = sheet.createRow(rowIndex++);
            writeCell(middleHeader, 0, "\u043f\u043e\n\u043f\u043e\u0440\u044f\u0434.", headerStyle);
            writeCell(middleHeader, 1, "\u043f\u043e\u0437. \u043f\u043e\n\u0441\u043c\u0435\u0442\u0435", headerStyle);
            writeCell(middleHeader, 2, "", headerStyle);
            writeCell(middleHeader, 3, "", headerStyle);
            writeCell(middleHeader, 4, "\u043a\u043e\u043b.", headerStyle);
            writeCell(middleHeader, 5, "\u0446\u0435\u043d\u0430", headerStyle);
            writeCell(middleHeader, 6, "\u0441\u0443\u043c\u043c\u0430", headerStyle);
            writeCell(middleHeader, 7, "\u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u043e \u0437\u0430\n\u043f\u0440\u0435\u0434. \u043f\u0435\u0440\u0438\u043e\u0434\u044b", headerStyle);
            mergeCellsIfNeeded(sheet, middleHeader.getRowNum(), 7, 8);
            writeCell(middleHeader, 9, "\u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u043e \u0437\u0430\n\u043e\u0442\u0447\u0435\u0442\u043d\u044b\u0439 \u043f\u0435\u0440\u0438\u043e\u0434", headerStyle);
            mergeCellsIfNeeded(sheet, middleHeader.getRowNum(), 9, 10);
            writeCell(middleHeader, 11, "\u0412\u0441\u0435\u0433\u043e \u0441 \u043d\u0430\u0447\u0430\u043b\u0430\n\u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u044f \u0440\u0430\u0431\u043e\u0442", headerStyle);
            mergeCellsIfNeeded(sheet, middleHeader.getRowNum(), 11, 12);
            writeCell(middleHeader, 13, "(%)", headerStyle);

            Row bottomHeader = sheet.createRow(rowIndex++);
            for (int c = 0; c <= lastColumn; c++) {
                writeCell(bottomHeader, c, "", headerStyle);
            }
            writeCell(bottomHeader, 7, "\u043a\u043e\u043b.", headerStyle);
            writeCell(bottomHeader, 8, "\u0441\u0443\u043c\u043c\u0430", headerStyle);
            writeCell(bottomHeader, 9, "\u043a\u043e\u043b.", headerStyle);
            writeCell(bottomHeader, 10, "\u0441\u0443\u043c\u043c\u0430", headerStyle);
            writeCell(bottomHeader, 11, "\u043a\u043e\u043b.", headerStyle);
            writeCell(bottomHeader, 12, "\u0441\u0443\u043c\u043c\u0430", headerStyle);

            Row indexHeader = sheet.createRow(rowIndex++);
            for (int c = 0; c <= lastColumn; c++) {
                writeCell(indexHeader, c, String.valueOf(c + 1), headerStyle);
            }

            for (Map<String, Object> row : rows) {
                Row workRow = sheet.createRow(rowIndex);
                writeSmartNumericCell(workRow, 0, toDouble(row.get("row_no")), integerStyle, decimalStyle);
                writeCell(workRow, 1, stringValue(row.get("estimate_no")), centerStyle);
                writeCell(workRow, 2, stringValue(row.get("work_name")), workNameStyle);
                writeCell(workRow, 3, stringValue(row.get("unit_name")), centerStyle);
                writeCell(workRow, 4, "", centerStyle);
                writeCell(workRow, 5, "", centerStyle);
                writeCell(workRow, 6, "", centerStyle);
                writeCell(workRow, 7, "", centerStyle);
                writeCell(workRow, 8, "", centerStyle);
                writeSmartNumericCell(workRow, 9, toDouble(row.get("period_quantity")), integerStyle, decimalStyle);
                writeSmartNumericCell(workRow, 10, toDouble(row.get("period_amount")), integerStyle, decimalStyle);
                writeCell(workRow, 11, "", centerStyle);
                writeCell(workRow, 12, "", centerStyle);
                writeCell(workRow, 13, formatPercent(toDouble(row.get("completion_percent"))), centerStyle);

                List<Map<String, Object>> materials = (List<Map<String, Object>>) row.getOrDefault("materials", List.of());
                int workRowIndex = rowIndex++;
                for (Map<String, Object> material : materials) {
                    Row materialRow = sheet.createRow(rowIndex++);
                    writeSmartNumericCell(materialRow, 0, toDouble(material.get("row_no")), integerStyle, decimalStyle);
                    writeCell(materialRow, 2, stringValue(material.get("material_name")), materialNameStyle);
                    writeCell(materialRow, 3, stringValue(material.get("unit_name")), centerStyle);
                    writeSmartNumericCell(materialRow, 9, toDouble(material.get("period_quantity")), integerStyle, decimalStyle);
                    for (int col : List.of(1,4,5,6,7,8,10,11,12,13)) {
                        if (materialRow.getCell(col) == null) {
                            writeCell(materialRow, col, "", centerStyle);
                        }
                    }
                }

                int span = Math.max(1, materials.size() + 1);
                if (!stringValue(row.get("estimate_no")).isBlank()) {
                    mergeCellsIfNeeded(sheet, workRowIndex, 1, rowIndex - 1, 1);
                }
            }

            rowIndex += 3;
            rowIndex = writeMergedValue(sheet, rowIndex, 1, 3, "\u041f\u0440\u043e\u0440\u0430\u0431 ____________________ " + firstNonBlank(header.get("foreman_name"), ""), boldLeftStyle);

            int[] widths = {8, 8, 52, 8, 10, 10, 10, 11, 11, 11, 11, 11, 11, 8};
            for (int i = 0; i < widths.length; i++) {
                sheet.setColumnWidth(i, widths[i] * 256);
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }

    private byte[] generateForm2Html(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        StringBuilder html = new StringBuilder();
        html.append("<!doctype html><html><head><meta charset='UTF-8'>");
        html.append("<style>");
        html.append("body{font-family:'Times New Roman',serif;margin:20px;color:#000;width:max-content;min-width:1200px;} ");
        html.append(".meta-table,.report-table{border-collapse:collapse;width:max-content;min-width:1200px;} ");
        html.append(".meta-table td,.report-table td,.report-table th{border:1px solid #000;padding:4px 6px;font-size:12px;vertical-align:middle;} ");
        html.append(".meta-table td{border:none;padding:2px 4px;} ");
        html.append(".italic{font-style:italic;} .center{text-align:center;} .right{text-align:right;} .bold{font-weight:700;} ");
        html.append(".title{font-weight:700;text-align:center;font-size:16px;margin-top:8px;} ");
        html.append(".subtitle{font-weight:700;text-align:center;font-size:14px;margin-bottom:8px;} ");
        html.append(".work-name{font-weight:700;} .material-name{font-style:italic;padding-left:18px;} ");
        html.append("</style></head><body>");

        html.append("<table class='meta-table'>");
        html.append("<tr><td class='bold' style='width:90px'>\u0417\u0430\u043a\u0430\u0437\u0447\u0438\u043a:</td><td class='italic' style='width:420px'>")
                .append(escapeHtml(stringValue(header.get("customer_name"))))
                .append("</td><td class='center bold' style='width:60px'>?</td><td class='center bold' style='width:110px'>\u0414\u0430\u0442\u0430</td><td class='center bold' colspan='2'>\u041e\u0442\u0447\u0435\u0442\u043d\u044b\u0439 \u043f\u0435\u0440\u0438\u043e\u0434</td></tr>");
        html.append("<tr><td class='bold'>\u041e\u0431\u044a\u0435\u043a\u0442:</td><td class='italic'>")
                .append(escapeHtml(stringValue(header.get("project_name")) + ", " + stringValue(header.get("project_address"))))
                .append("</td><td class='center bold'>\u0434\u043e\u043a\u0443\u043c\u0435\u043d\u0442\u0430</td><td class='center bold'>\u0441\u043e\u0441\u0442\u0430\u0432\u043b\u0435\u043d\u0438\u044f</td><td class='center bold'>\u0441</td><td class='center bold'>\u043f\u043e</td></tr>");
        html.append("<tr><td class='bold'>\u0412\u0438\u0434 \u0421\u041c\u0420:</td><td class='italic'>")
                .append(escapeHtml(stringValue(header.get("construction_type"))))
                .append("</td><td class='center'>")
                .append(escapeHtml(stringValue(header.get("report_number"))))
                .append("</td><td class='center'>")
                .append(escapeHtml(stringValue(header.get("report_date_label"))))
                .append("</td><td class='center'>")
                .append(escapeHtml(stringValue(header.get("period_from_label"))))
                .append("</td><td class='center'>")
                .append(escapeHtml(stringValue(header.get("period_to_label"))))
                .append("</td></tr></table>");

        html.append("<div class='title'>\u0410\u041a\u0422</div>");
        html.append("<div class='subtitle'>\u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u043d\u044b\u0445 \u0440\u0430\u0431\u043e\u0442 \u0437\u0430 ").append(escapeHtml(stringValue(header.get("period_label")))).append("</div>");

        html.append("<table class='report-table'>");
        html.append("<colgroup>");
        html.append("<col style='width:46px'>");
        html.append("<col style='width:58px'>");
        html.append("<col style='width:360px'>");
        html.append("<col style='width:56px'>");
        html.append("<col style='width:64px'>");
        html.append("<col style='width:74px'>");
        html.append("<col style='width:86px'>");
        html.append("<col style='width:78px'>");
        html.append("<col style='width:86px'>");
        html.append("<col style='width:78px'>");
        html.append("<col style='width:86px'>");
        html.append("<col style='width:78px'>");
        html.append("<col style='width:86px'>");
        html.append("<col style='width:62px'>");
        html.append("</colgroup><thead>");
        html.append("<tr>");
        html.append("<th colspan='2'>\u041d\u043e\u043c\u0435\u0440</th><th rowspan='3'>\u041d\u0430\u0438\u043c\u0435\u043d\u043e\u0432\u0430\u043d\u0438\u0435 \u0440\u0430\u0431\u043e\u0442</th><th rowspan='3'>\u0435\u0434.<br>\u0438\u0437\u043c.</th>");
        html.append("<th colspan='3'>\u043f\u043e \u0434\u043e\u0433\u043e\u0432\u043e\u0440\u0443</th><th colspan='7'>\u0424\u0430\u043a\u0442\u0438\u0447\u0435\u0441\u043a\u043e\u0435 \u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u0435 \u0444\u0438\u0437\u0438\u0447\u0435\u0441\u043a\u0438\u0445 \u043e\u0431\u044a\u0435\u043c\u043e\u0432 \u0440\u0430\u0431\u043e\u0442 \u0438 \u0441\u0442\u043e\u0438\u043c\u043e\u0441\u0442\u0438</th>");
        html.append("</tr>");
        html.append("<tr>");
        html.append("<th>\u043f\u043e<br>\u043f\u043e\u0440\u044f\u0434.</th><th>\u043f\u043e\u0437. \u043f\u043e<br>\u0441\u043c\u0435\u0442\u0435</th>");
        html.append("<th>\u043a\u043e\u043b.</th><th>\u0446\u0435\u043d\u0430</th><th>\u0441\u0443\u043c\u043c\u0430</th>");
        html.append("<th colspan='2'>\u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u043e \u0437\u0430<br>\u043f\u0440\u0435\u0434. \u043f\u0435\u0440\u0438\u043e\u0434\u044b</th>");
        html.append("<th colspan='2'>\u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u043e \u0437\u0430<br>\u043e\u0442\u0447\u0435\u0442\u043d\u044b\u0439 \u043f\u0435\u0440\u0438\u043e\u0434</th>");
        html.append("<th colspan='2'>\u0412\u0441\u0435\u0433\u043e \u0441 \u043d\u0430\u0447\u0430\u043b\u0430<br>\u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u044f \u0440\u0430\u0431\u043e\u0442</th>");
        html.append("<th>(%)</th>");
        html.append("</tr>");
        html.append("<tr>");
        html.append("<th>1</th><th>2</th><th>5</th><th>6</th><th>7</th><th>8</th><th>9</th><th>10</th><th>11</th><th>12</th><th>13</th><th>14</th>");
        html.append("</tr></thead><tbody>");

        for (Map<String, Object> row : rows) {
            List<Map<String, Object>> materials = (List<Map<String, Object>>) row.getOrDefault("materials", List.of());
            int span = Math.max(1, materials.size() + 1);

            html.append("<tr>");
            html.append("<td class='center'>").append(formatSmartNumber(toDouble(row.get("row_no")))).append("</td>");
            html.append("<td class='center' rowspan='").append(span).append("'>").append(escapeHtml(stringValue(row.get("estimate_no")))).append("</td>");
            html.append("<td class='work-name'>").append(escapeHtml(stringValue(row.get("work_name")))).append("</td>");
            html.append("<td class='center'>").append(escapeHtml(stringValue(row.get("unit_name")))).append("</td>");
            html.append("<td class='center'></td>");
            html.append("<td class='center'></td>");
            html.append("<td class='center'></td>");
            html.append("<td class='center'></td>");
            html.append("<td class='center'></td>");
            html.append("<td class='center'>").append(formatSmartNumber(toDouble(row.get("period_quantity")))).append("</td>");
            html.append("<td class='center'>").append(formatSmartNumber(toDouble(row.get("period_amount")))).append("</td>");
            html.append("<td class='center'></td>");
            html.append("<td class='center'></td>");
            html.append("<td class='center'>").append(escapeHtml(formatPercent(toDouble(row.get("completion_percent"))))).append("</td>");
            html.append("</tr>");

            for (Map<String, Object> material : materials) {
                html.append("<tr>");
                html.append("<td class='center'>").append(formatSmartNumber(toDouble(material.get("row_no")))).append("</td>");
                html.append("<td class='material-name'>").append(escapeHtml(stringValue(material.get("material_name")))).append("</td>");
                html.append("<td class='center'>").append(escapeHtml(stringValue(material.get("unit_name")))).append("</td>");
                html.append("<td></td><td></td><td></td><td></td><td></td>");
                html.append("<td class='center'>").append(formatSmartNumber(toDouble(material.get("period_quantity")))).append("</td>");
                html.append("<td></td><td></td><td></td><td></td>");
                html.append("</tr>");
            }
        }

        html.append("</tbody></table>");
        html.append("<div style='margin-top:80px;font-weight:700;'>\u041f\u0440\u043e\u0440\u0430\u0431 ____________________</div>");
        html.append("</body></html>");
        return html.toString().getBytes(StandardCharsets.UTF_8);
    }

    private List<Map<String, Object>> buildForm2PrintableRows(List<Map<String, Object>> rows) {
        List<Map<String, Object>> printableRows = new ArrayList<>();

        for (Map<String, Object> row : rows) {
            Map<String, Object> workRow = new LinkedHashMap<>();
            workRow.put("row_no", formatSmartNumber(toDouble(row.get("row_no"))));
            workRow.put("estimate_no", stringValue(row.get("estimate_no")));
            workRow.put("item_name", stringValue(row.get("work_name")));
            workRow.put("unit_name", stringValue(row.get("unit_name")));
            workRow.put("contract_quantity", null);
            workRow.put("contract_price", null);
            workRow.put("contract_amount", null);
            workRow.put("previous_quantity", null);
            workRow.put("previous_amount", null);
            workRow.put("period_quantity", nullableDouble(row.get("period_quantity")));
            workRow.put("period_amount", nullableDouble(row.get("period_amount")));
            workRow.put("total_quantity", null);
            workRow.put("total_amount", null);
            workRow.put("completion_percent_label", formatPercent(toDouble(row.get("completion_percent"))));
            workRow.put("item_type", "work");
            printableRows.add(workRow);

            List<Map<String, Object>> materials = (List<Map<String, Object>>) row.getOrDefault("materials", List.of());
            for (Map<String, Object> material : materials) {
                Map<String, Object> materialRow = new LinkedHashMap<>();
                materialRow.put("row_no", formatSmartNumber(toDouble(material.get("row_no"))));
                materialRow.put("estimate_no", "");
                materialRow.put("item_name", "   " + stringValue(material.get("material_name")));
                materialRow.put("unit_name", stringValue(material.get("unit_name")));
                materialRow.put("contract_quantity", null);
                materialRow.put("contract_price", null);
                materialRow.put("contract_amount", null);
                materialRow.put("previous_quantity", null);
                materialRow.put("previous_amount", null);
                materialRow.put("period_quantity", nullableDouble(material.get("period_quantity")));
                materialRow.put("period_amount", null);
                materialRow.put("total_quantity", null);
                materialRow.put("total_amount", null);
                materialRow.put("completion_percent_label", "");
                materialRow.put("item_type", "material");
                printableRows.add(materialRow);
            }
        }

        return printableRows;
    }

    private Double nullableDouble(Object value) {
        if (value == null) {
            return null;
        }
        return toDouble(value);
    }


    private double toDouble(Object value) {
        if (value == null) {
            return 0;
        }
        if (value instanceof Number number) {
            return number.doubleValue();
        }
        if (value instanceof String str) {
            String normalized = str.trim().replace(" ", "").replace(",", ".");
            if (normalized.isBlank()) {
                return 0;
            }
            try {
                return Double.parseDouble(normalized);
            } catch (NumberFormatException ignored) {
                return 0;
            }
        }
        return 0;
    }

    private Object firstNonNull(Object primary, Object fallback) {
        return primary != null ? primary : fallback;
    }

    private String stringValue(Object value) {
        return value == null ? "" : String.valueOf(value);
    }

    private String firstNonBlank(Object primary, Object fallback) {
        String primaryValue = stringValue(primary);
        if (!primaryValue.isBlank()) {
            return primaryValue;
        }
        return stringValue(fallback);
    }

    private String projectAddressSuffix(Map header) {
        String address = stringValue(header.get("project_address"));
        return address.isBlank() ? "" : " по ул. " + address;
    }

    private String formatCreatedAt(Object value) {
        String raw = stringValue(value);
        if (raw.isBlank()) {
            return "";
        }

        try {
            String normalized = raw.substring(0, Math.min(raw.length(), 19)).replace(' ', 'T');
            LocalDateTime dateTime = LocalDateTime.parse(normalized, DateTimeFormatter.ISO_LOCAL_DATE_TIME);
            String[] months = {
                    "января", "февраля", "марта", "апреля", "мая", "июня",
                    "июля", "августа", "сентября", "октября", "ноября", "декабря"
            };
            return String.format("\"%02d\" %s %d г.", dateTime.getDayOfMonth(), months[dateTime.getMonthValue() - 1], dateTime.getYear());
        } catch (RuntimeException ignored) {
            return raw;
        }
    }

    private String amountToWords(double amount) {
        BigDecimal rounded = BigDecimal.valueOf(amount).setScale(2, RoundingMode.HALF_UP);
        long soms = rounded.longValue();
        int tyiyn = rounded.remainder(BigDecimal.ONE).movePointRight(2).abs().intValue();

        return capitalize(numberToWords(soms)) + " сом " + String.format("%02d", tyiyn) + " тыйын";
    }

    private String numberToWords(long value) {
        if (value == 0) {
            return "ноль";
        }

        StringBuilder result = new StringBuilder();
        appendGroup(result, (int) (value / 1_000_000_000), "", "миллиард", "миллиарда", "миллиардов", false);
        appendGroup(result, (int) ((value / 1_000_000) % 1000), "", "миллион", "миллиона", "миллионов", false);
        appendGroup(result, (int) ((value / 1000) % 1000), "", "тысяча", "тысячи", "тысяч", true);
        appendGroup(result, (int) (value % 1000), "", "", "", "", false);

        return result.toString().trim().replaceAll("\\s+", " ");
    }

    private void appendGroup(
            StringBuilder result,
            int value,
            String zeroForm,
            String oneForm,
            String twoForm,
            String manyForm,
            boolean feminine
    ) {
        if (value == 0) {
            if (!zeroForm.isBlank()) {
                appendWord(result, zeroForm);
            }
            return;
        }

        int hundreds = value / 100;
        int tensUnits = value % 100;
        int tens = tensUnits / 10;
        int units = tensUnits % 10;

        String[] hundredsWords = {"", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"};
        String[] tensWords = {"", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"};
        String[] teensWords = {"десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"};
        String[] unitsMasculine = {"", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"};
        String[] unitsFeminine = {"", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"};

        appendWord(result, hundredsWords[hundreds]);

        if (tensUnits >= 10 && tensUnits <= 19) {
            appendWord(result, teensWords[tensUnits - 10]);
        } else {
            appendWord(result, tensWords[tens]);
            appendWord(result, feminine ? unitsFeminine[units] : unitsMasculine[units]);
        }

        if (!oneForm.isBlank()) {
            appendWord(result, pluralForm(value, oneForm, twoForm, manyForm));
        }
    }

    private String pluralForm(int value, String oneForm, String twoForm, String manyForm) {
        int lastTwo = value % 100;
        int last = value % 10;
        if (lastTwo >= 11 && lastTwo <= 14) {
            return manyForm;
        }
        if (last == 1) {
            return oneForm;
        }
        if (last >= 2 && last <= 4) {
            return twoForm;
        }
        return manyForm;
    }

    private void appendWord(StringBuilder result, String word) {
        if (word == null || word.isBlank()) {
            return;
        }
        if (!result.isEmpty()) {
            result.append(' ');
        }
        result.append(word);
    }

    private String capitalize(String value) {
        if (value == null || value.isBlank()) {
            return "";
        }
        return value.substring(0, 1).toUpperCase() + value.substring(1);
    }

    private void appendHtmlTotalsRow(
            StringBuilder html,
            List<Map<String, Object>> materials,
            Map<Integer, Map<String, Object>> totalsByMaterial,
            String label,
            String field
    ) {
        html.append("<tr>");
        html.append("<td class=\"label\">").append(escapeHtml(label)).append("</td>");
        html.append("<td class=\"label\"></td>");
        html.append("<td class=\"label\"></td>");
        for (Map<String, Object> material : materials) {
            Map<String, Object> total = totalsByMaterial.get(toInt(material.get("material_id")));
            double value = 0;
            if (total != null) {
                if ("deviation_quantity".equals(field)) {
                    value = toDouble(total.get(field));
                } else {
                    value = toDouble(total.get(field));
                }
            }
            String materialClass = materialCssClass(material);
            html.append("<td class=\"num ").append(materialClass).append("\">Х</td>");
            html.append("<td class=\"num ").append(materialClass).append("\">")
                    .append("deviation_quantity".equals(field) ? formatSignedSmartNumber(value) : formatSmartNumber(value))
                    .append("</td>");
        }
        html.append("</tr>");
    }

    private String materialCssClass(Map<String, Object> material) {
        boolean hasFact = Boolean.TRUE.equals(material.get("has_fact"));
        boolean isAdditional = Boolean.TRUE.equals(material.get("is_additional"));
        boolean fromEstimate = Boolean.TRUE.equals(material.get("from_estimate"));

        if (hasFact && isAdditional) {
            return "additional";
        }
        if (!hasFact && fromEstimate) {
            return "estimate";
        }
        return "base";
    }

    private String formatSmartNumber(double value) {
        if (Math.abs(value - Math.rint(value)) < 0.0000001d) {
            return String.valueOf((long) Math.rint(value));
        }
        BigDecimal decimal = BigDecimal.valueOf(value).stripTrailingZeros();
        return decimal.toPlainString();
    }

    private String formatSignedSmartNumber(double value) {
        String formatted = formatSmartNumber(value);
        if (Math.abs(value) < 0.0000001d) {
            return "0";
        }
        return value > 0 ? "+" + formatted : formatted;
    }

    private String formatPercent(double value) {
        if (value <= 0) {
            return "";
        }
        return formatSmartNumber(value) + "%";
    }

    private String normalizeFormat(String format, String fallback) {
        if (format == null || format.isBlank()) {
            return fallback;
        }
        String normalized = format.toLowerCase();
        return switch (normalized) {
            case "pdf", "html", "xlsx", "docx" -> normalized;
            default -> fallback;
        };
    }

    private String buildForm29SectionTitle(Map<String, Object> section) {
        String sectionTitle = stringValue(section.get("stage_name"));
        String subsectionName = stringValue(section.get("subsection_name"));
        if (!subsectionName.isBlank()) {
            return sectionTitle + " / " + subsectionName;
        }
        return sectionTitle;
    }

    private String form29Bucket(int sectionOrder, int rowOrder, String type) {
        return String.format("%03d|%05d|%s", sectionOrder, rowOrder, type);
    }

    private Map<String, Object> buildForm29BodyRow(
            String rowBucket,
            String rowLabel,
            String rowType,
            Map<String, Object> material,
            int materialOrder,
            String metricBucket,
            String metricLabel,
            String displayValue
    ) {
        Map<String, Object> row = new HashMap<>();
        row.put("row_bucket", rowBucket + "|" + rowLabel);
        row.put("row_label", rowLabel);
        row.put("row_type", rowType);
        row.put("material_bucket", String.format("%03d|%s", materialOrder, buildForm29MaterialLabel(material)));
        row.put("material_label", buildForm29MaterialLabel(material));
        row.put("metric_bucket", metricBucket + "|" + metricLabel);
        row.put("metric_label", metricLabel);
        row.put("display_value", displayValue);
        return row;
    }

    private Map<String, Object> buildForm29TotalRow(
            String totalBucket,
            String totalLabel,
            Map<String, Object> material,
            int materialOrder,
            String metricBucket,
            String metricLabel,
            String displayValue
    ) {
        Map<String, Object> row = new HashMap<>();
        row.put("total_bucket", totalBucket);
        row.put("total_label", totalLabel);
        row.put("material_bucket", String.format("%03d|%s", materialOrder, buildForm29MaterialLabel(material)));
        row.put("material_label", buildForm29MaterialLabel(material));
        row.put("metric_bucket", metricBucket + "|" + metricLabel);
        row.put("metric_label", metricLabel);
        row.put("display_value", displayValue);
        return row;
    }

    private String buildForm29MaterialLabel(Map<String, Object> material) {
        String materialName = stringValue(material.get("material_name"));
        String unitName = stringValue(material.get("unit_name"));
        if (unitName.isBlank()) {
            return materialName;
        }
        return materialName + "\n" + unitName;
    }

    private String formatForm29Norm(Map<String, Object> cell) {
        if (cell == null) {
            return "";
        }
        Object normValue = cell.get("norm_value");
        if (normValue instanceof Number) {
            return formatSmartNumber(toDouble(normValue));
        }
        return stringValue(normValue);
    }

    public byte[] generateSchedule(ReportData reportData, String format) throws Exception {
        if ("xlsx".equalsIgnoreCase(format)) {
            return generateProjectScheduleXlsx(reportData);
        }

        if ("html".equalsIgnoreCase(format)) {
            return generateScheduleHtml(reportData);
        }

        BufferedImage image = renderProjectScheduleImage(reportData);
        InputStream jrxmlStream = new ClassPathResource("reports/ScheduleImageReport.jrxml").getInputStream();
        JasperReport report = JasperCompileManager.compileReport(jrxmlStream);

        Map<String, Object> params = new HashMap<>();
        params.put("reportImage", (Image) image);

        JasperPrint print = JasperFillManager.fillReport(report, params, new JREmptyDataSource(1));
        return switch (format == null ? "pdf" : format.toLowerCase()) {
            case "docx" -> exportDocx(print);
            default -> JasperExportManager.exportReportToPdf(print);
        };
    }

    public byte[] generateMaterialSchedule(ReportData reportData, String format) throws Exception {
        return generateSchedule(reportData, format);
    }

    public byte[] generateForm19(ReportData reportData, String format) throws Exception {
        if ("xlsx".equalsIgnoreCase(format)) {
            return generateForm19XlsxDetailed(reportData);
        }

        if ("html".equalsIgnoreCase(format)) {
            return generateForm19Html(reportData);
        }

        BufferedImage image = renderForm19ImageDetailed(reportData);
        InputStream jrxmlStream = new ClassPathResource("reports/ScheduleImageReport.jrxml").getInputStream();
        JasperReport report = JasperCompileManager.compileReport(jrxmlStream);

        Map<String, Object> params = new HashMap<>();
        params.put("reportImage", (Image) image);

        JasperPrint print = JasperFillManager.fillReport(report, params, new JREmptyDataSource(1));
        return switch (format == null ? "pdf" : format.toLowerCase()) {
            case "docx" -> exportDocx(print);
            default -> JasperExportManager.exportReportToPdf(print);
        };
    }

    public byte[] generateMbpWriteOff(ReportData reportData, String format) throws Exception {
        if ("xlsx".equalsIgnoreCase(format)) {
            return generateMbpWriteOffXlsx(reportData);
        }

        if ("html".equalsIgnoreCase(format)) {
            return generateMbpWriteOffHtml(reportData);
        }

        BufferedImage image = renderMbpWriteOffImage(reportData);
        InputStream jrxmlStream = new ClassPathResource("reports/ScheduleImageReport.jrxml").getInputStream();
        JasperReport report = JasperCompileManager.compileReport(jrxmlStream);

        Map<String, Object> params = new HashMap<>();
        params.put("reportImage", (Image) image);

        JasperPrint print = JasperFillManager.fillReport(report, params, new JREmptyDataSource(1));
        return switch (format == null ? "pdf" : format.toLowerCase()) {
            case "docx" -> exportDocx(print);
            default -> JasperExportManager.exportReportToPdf(print);
        };
    }

    public byte[] generateProjectsOverview(ReportData reportData, String format) throws Exception {
        if ("xlsx".equalsIgnoreCase(format)) {
            return generateProjectsOverviewXlsx(reportData);
        }

        if ("html".equalsIgnoreCase(format)) {
            return generateProjectsOverviewHtml(reportData);
        }

        BufferedImage image = renderProjectsOverviewImage(reportData);
        InputStream jrxmlStream = new ClassPathResource("reports/ScheduleImageReport.jrxml").getInputStream();
        JasperReport report = JasperCompileManager.compileReport(jrxmlStream);

        Map<String, Object> params = new HashMap<>();
        params.put("reportImage", (Image) image);

        JasperPrint print = JasperFillManager.fillReport(report, params, new JREmptyDataSource(1));
        return switch (format == null ? "pdf" : format.toLowerCase()) {
            case "docx" -> exportDocx(print);
            default -> JasperExportManager.exportReportToPdf(print);
        };
    }

    private byte[] generateScheduleHtml(ReportData reportData) throws Exception {
        ScheduleRenderModel model = buildProjectScheduleModel(reportData);
        StringBuilder html = new StringBuilder();

        html.append("<!doctype html><html><head><meta charset=\"UTF-8\">");
        html.append("<style>");
        html.append("html,body{margin:0;padding:0;background:#f3f4f6;color:#000;} ");
        html.append("body{font-family:Arial,sans-serif;padding:16px;width:max-content;min-width:560px;box-sizing:border-box;} ");
        html.append(".page{width:max-content;background:#fff;padding:20px;box-shadow:0 2px 10px rgba(0,0,0,.08);} ");
        html.append(".meta{font-size:12px;font-weight:700;margin-bottom:4px;} ");
        html.append(".title{font-size:20px;font-weight:700;text-align:center;margin:28px 0 6px;white-space:nowrap;} ");
        html.append(".subtitle{font-size:14px;font-weight:700;text-align:center;margin-bottom:18px;white-space:nowrap;} ");
        html.append("table{border-collapse:collapse;width:max-content;} ");
        html.append("th,td{border:1px solid #000;padding:0 4px;font-size:12px;height:24px;vertical-align:middle;} ");
        html.append("th{font-weight:700;text-align:center;} ");
        html.append(".num{width:42px;text-align:center;} .name{width:310px;text-align:left;} ");
        html.append(".name-head{text-align:center;} .slot{width:22px;min-width:22px;text-align:center;padding:0;} ");
        html.append(".filled{background:#5b8fd1;} .section td{background:#f4f4f5;font-weight:700;} ");
        html.append(".footer{font-weight:700;margin-top:22px;font-size:12px;} ");
        html.append("</style></head><body><div class='page'>");

        html.append("<div class='meta'>Заказчик: ")
                .append(escapeHtml(model.customerName()))
                .append("</div>");
        html.append("<div class='meta'>Объект: ")
                .append(escapeHtml(model.projectName()))
                .append("</div>");
        html.append("<div class='title'>")
                .append(escapeHtml(model.reportTitle()))
                .append("</div>");
        html.append("<div class='subtitle'>по объекту: ")
                .append(escapeHtml(model.projectName()))
                .append("</div>");

        html.append("<table><thead><tr>");
        html.append("<th class='num' rowspan='3'>№</th>");
        html.append("<th class='name name-head' rowspan='3'>")
                .append(escapeHtml(model.nameColumn()))
                .append("</th>");

        appendScheduleGroupHeader(html, model.slots(), true);
        html.append("</tr><tr>");
        appendScheduleGroupHeader(html, model.slots(), false);
        html.append("</tr><tr>");
        for (ScheduleSlot slot : model.slots()) {
            html.append("<th class='slot'>")
                    .append(escapeHtml(slot.slotLabel()))
                    .append("</th>");
        }
        html.append("</tr></thead><tbody>");

        for (ScheduleRowData rowData : model.rows()) {
            if (rowData.rowNo() == 0) {
                html.append("<tr class='section'><td class='num'></td><td class='name'>")
                        .append(escapeHtml(rowData.label()))
                        .append("</td>");
                for (int i = 0; i < model.slots().size(); i++) {
                    html.append("<td class='slot'></td>");
                }
                html.append("</tr>");
                continue;
            }

            html.append("<tr><td class='num'>")
                    .append(rowData.rowNo())
                    .append("</td><td class='name'>")
                    .append(escapeHtml(rowData.label()))
                    .append("</td>");
            for (int i = 0; i < model.slots().size(); i++) {
                boolean filled = i < rowData.filled().size() && Boolean.TRUE.equals(rowData.filled().get(i));
                html.append("<td class='slot")
                        .append(filled ? " filled" : "")
                        .append("'></td>");
            }
            html.append("</tr>");
        }

        html.append("</tbody></table>");
        html.append("<div class='footer'>Главный инженер: ____________________</div>");
        html.append("<div class='footer'>Инженер ПТО: ____________________</div>");
        html.append("</div></body></html>");

        return html.toString().getBytes(StandardCharsets.UTF_8);
    }

    private void appendScheduleGroupHeader(StringBuilder html, List<ScheduleSlot> slots, boolean yearLevel) {
        int index = 0;
        while (index < slots.size()) {
            ScheduleSlot current = slots.get(index);
            String label = yearLevel ? current.yearLabel() : current.monthLabel();
            int colspan = 1;

            while (index + colspan < slots.size()) {
                ScheduleSlot next = slots.get(index + colspan);
                String nextLabel = yearLevel ? next.yearLabel() : next.monthLabel();
                if (!label.equals(nextLabel)) {
                    break;
                }
                colspan++;
            }

            html.append("<th class='slot' colspan='")
                    .append(colspan)
                    .append("'>")
                    .append(escapeHtml(label))
                    .append("</th>");
            index += colspan;
        }
    }

    private byte[] generateForm19Html(ReportData reportData) throws Exception {
        BufferedImage image = renderForm19ImageDetailed(reportData);
        ByteArrayOutputStream imageOut = new ByteArrayOutputStream();
        javax.imageio.ImageIO.write(image, "png", imageOut);
        String base64 = Base64.getEncoder().encodeToString(imageOut.toByteArray());
        return buildImagePreviewHtml(base64, "form19");
    }

    private byte[] generateMbpWriteOffHtml(ReportData reportData) throws Exception {
        BufferedImage image = renderMbpWriteOffImage(reportData);
        ByteArrayOutputStream imageOut = new ByteArrayOutputStream();
        javax.imageio.ImageIO.write(image, "png", imageOut);
        String base64 = Base64.getEncoder().encodeToString(imageOut.toByteArray());
        return buildImagePreviewHtml(base64, "mbp-write-off");
    }

    private byte[] generateProjectsOverviewHtml(ReportData reportData) throws Exception {
        BufferedImage image = renderProjectsOverviewImage(reportData);
        ByteArrayOutputStream imageOut = new ByteArrayOutputStream();
        javax.imageio.ImageIO.write(image, "png", imageOut);
        String base64 = Base64.getEncoder().encodeToString(imageOut.toByteArray());
        return buildImagePreviewHtml(base64, "projects-overview");
    }

    private byte[] generateMbpWriteOffXlsx(ReportData reportData) throws Exception {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Списание МБП");

            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, null);
            CellStyle approvalStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle integerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);

            headerStyle.setWrapText(true);
            integerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));

            int[] widths = {1600, 10500, 1800, 2200, 3200, 4200};
            for (int i = 0; i < widths.length; i++) {
                sheet.setColumnWidth(i, widths[i]);
            }

            int rowIndex = 0;
            rowIndex = writeMergedValue(sheet, rowIndex, 4, 5, "\"УТВЕРЖДАЮ\"", approvalStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 4, 5, "Генеральный директор", approvalStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 4, 5, "ОсОО \"СК Дрим Хаус\"", approvalStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 4, 5, "Темиркулов Ш. Т.", approvalStyle);
            rowIndex++;

            rowIndex = writeMergedValue(sheet, rowIndex, 0, 5, "АКТ", titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 5, "на списание малоценных и быстроизнашиваемых материалов", titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 5, "за " + stringValue(header.get("period_label")), titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 5, "Объект: " + stringValue(header.get("project_name")) + projectAddressSuffix(header), metaStyle);

            Row tableHeader = sheet.createRow(rowIndex++);
            tableHeader.setHeightInPoints(30);
            writeCell(tableHeader, 0, "№\nп/п", headerStyle);
            writeCell(tableHeader, 1, "Наименование материала", headerStyle);
            writeCell(tableHeader, 2, "Ед.\nизм.", headerStyle);
            writeCell(tableHeader, 3, "Кол-\nво", headerStyle);
            writeCell(tableHeader, 4, "Дата\nсписания", headerStyle);
            writeCell(tableHeader, 5, "Примечание", headerStyle);

            for (Map<String, Object> rowData : rows) {
                Row row = sheet.createRow(rowIndex++);
                writeSmartNumericCell(row, 0, toDouble(rowData.get("row_no")), integerStyle, decimalStyle);
                writeCell(row, 1, stringValue(rowData.get("material_name")), textStyle);
                writeCell(row, 2, stringValue(rowData.get("unit_name")), centerStyle);
                writeSmartNumericCell(row, 3, toDouble(rowData.get("quantity")), integerStyle, decimalStyle);
                writeCell(row, 4, stringValue(rowData.get("posted_at_label")), centerStyle);
                writeCell(row, 5, stringValue(rowData.get("note")), textStyle);
            }

            rowIndex += 2;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 3, "Прораб : __________________________", metaStyle);

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private byte[] generateProjectsOverviewXlsx(ReportData reportData) throws Exception {
        Map data = reportData.data();
        Map header = reportData.header();
        Map totals = (Map) data.getOrDefault("totals", Map.of());
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Проекты");

            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, null);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle integerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle moneyIntegerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle moneyDecimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);

            headerStyle.setWrapText(true);
            integerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.##"));
            moneyIntegerStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            moneyDecimalStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.##"));

            int[] widths = {1600, 8000, 3600, 5200, 7000, 3400, 3400, 4200, 4200, 3600, 3200, 3200, 3600, 3600, 3200, 3200, 3200, 3200, 3200};
            for (int i = 0; i < widths.length; i++) {
                sheet.setColumnWidth(i, widths[i]);
            }

            int lastColumn = widths.length - 1;
            int rowIndex = 0;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "СВОДНЫЙ ОТЧЕТ ПО ПРОЕКТАМ", titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "Дата формирования: " + formatIsoDate(header.get("report_date")), metaStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "Проектов: " + formatSmartNumber(toDouble(header.get("projects_count"))), metaStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "План бюджет: " + formatSmartNumber(toDouble(totals.get("total_planned_budget"))) + " | Факт бюджет: " + formatSmartNumber(toDouble(totals.get("total_actual_budget"))) + " | Средний прогресс: " + formatSmartNumber(toDouble(totals.get("avg_progress_percent"))) + "%", metaStyle);
            rowIndex++;

            Row headerRow = sheet.createRow(rowIndex++);
            String[] headers = {
                    "№", "Проект", "Статус", "Заказчик", "Адрес", "Начало", "Окончание",
                    "План бюджет", "Факт бюджет", "% бюджета", "Прогресс %",
                    "Блоки", "Площ. общ.", "Площ. продаж",
                    "Заявки", "Закупы", "АВР", "Списания", "Перемещения"
            };
            for (int i = 0; i < headers.length; i++) {
                writeCell(headerRow, i, headers[i], headerStyle);
            }

            int index = 1;
            for (Map<String, Object> rowData : rows) {
                Row row = sheet.createRow(rowIndex++);
                writeSmartNumericCell(row, 0, index++, integerStyle, decimalStyle);
                writeCell(row, 1, stringValue(rowData.get("project_name")), textStyle);
                writeCell(row, 2, stringValue(rowData.get("status_name")), centerStyle);
                writeCell(row, 3, stringValue(rowData.get("customer_name")), textStyle);
                writeCell(row, 4, stringValue(rowData.get("address")), textStyle);
                writeCell(row, 5, formatIsoDate(rowData.get("start_date")), centerStyle);
                writeCell(row, 6, formatIsoDate(rowData.get("end_date")), centerStyle);
                writeSmartNumericCell(row, 7, toDouble(rowData.get("planned_budget")), moneyIntegerStyle, moneyDecimalStyle);
                writeSmartNumericCell(row, 8, toDouble(rowData.get("actual_budget")), moneyIntegerStyle, moneyDecimalStyle);
                writeSmartNumericCell(row, 9, toDouble(rowData.get("budget_percent")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 10, toDouble(rowData.get("progress_percent")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 11, toDouble(rowData.get("blocks_count")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 12, toDouble(rowData.get("total_area")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 13, toDouble(rowData.get("sale_area")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 14, toDouble(rowData.get("material_requests_count")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 15, toDouble(rowData.get("purchase_orders_count")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 16, toDouble(rowData.get("signed_work_performed_count")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 17,
                        toDouble(rowData.get("material_write_off_count")) +
                        toDouble(rowData.get("mbp_write_off_count")) +
                        toDouble(rowData.get("processing_write_off_count")),
                        integerStyle, decimalStyle
                );
                writeSmartNumericCell(row, 18, toDouble(rowData.get("transfer_count")), integerStyle, decimalStyle);
            }

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private BufferedImage renderProjectsOverviewImage(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        Map totals = (Map) data.getOrDefault("totals", Map.of());
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        int margin = 24;
        int[] columnWidths = {42, 200, 90, 155, 180, 76, 76, 96, 96, 76, 76, 58, 72, 72, 58, 58, 58, 66, 66};
        int tableWidth = 0;
        for (int width : columnWidths) {
            tableWidth += width;
        }

        int rowHeight = 26;
        int headerHeight = 48;
        int topHeight = 170;
        int footerHeight = 40;
        int width = margin * 2 + tableWidth;
        int height = Math.max(560, topHeight + headerHeight + Math.max(1, rows.size()) * rowHeight + footerHeight);

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(Color.WHITE);
        g.fillRect(0, 0, width, height);

        java.awt.Font regular = new java.awt.Font("Arial", java.awt.Font.PLAIN, 11);
        java.awt.Font bold = new java.awt.Font("Arial", java.awt.Font.BOLD, 11);
        java.awt.Font title = new java.awt.Font("Arial", java.awt.Font.BOLD, 18);

        g.setColor(Color.BLACK);
        g.setFont(title);
        FontMetrics titleMetrics = g.getFontMetrics();
        drawCenteredString(g, "СВОДНЫЙ ОТЧЕТ ПО ПРОЕКТАМ", margin, 28, tableWidth, 24, titleMetrics);

        g.setFont(bold);
        FontMetrics boldMetrics = g.getFontMetrics();
        int y = 62;
        drawLeftString(g, "Дата формирования: " + formatIsoDate(header.get("report_date")), margin, y, tableWidth, 18, boldMetrics);
        y += 18;
        drawLeftString(g, "Проектов: " + formatSmartNumber(toDouble(header.get("projects_count"))), margin, y, tableWidth, 18, boldMetrics);
        y += 18;
        drawLeftString(
                g,
                "План бюджет: " + formatSmartNumber(toDouble(totals.get("total_planned_budget")))
                        + "   Факт бюджет: " + formatSmartNumber(toDouble(totals.get("total_actual_budget")))
                        + "   Остаток: " + formatSmartNumber(toDouble(totals.get("total_remaining_budget"))),
                margin,
                y,
                tableWidth,
                18,
                boldMetrics
        );
        y += 18;
        drawLeftString(
                g,
                "Средний прогресс: " + formatSmartNumber(toDouble(totals.get("avg_progress_percent"))) + "%"
                        + "   Блоков: " + formatSmartNumber(toDouble(totals.get("total_blocks")))
                        + "   Складов: " + formatSmartNumber(toDouble(totals.get("total_warehouses")))
                        + "   Смет: " + formatSmartNumber(toDouble(totals.get("total_estimates"))),
                margin,
                y,
                tableWidth,
                18,
                boldMetrics
        );

        String[] headers = {
                "№", "Проект", "Статус", "Заказчик", "Адрес", "Начало", "Окончание",
                "План\nбюджет", "Факт\nбюджет", "%\nбюдж.", "Прогр.\n%", "Блоки",
                "Общ.\nплощ.", "Прод.\nплощ.", "Заявки", "Закупы", "АВР", "Спис.", "Перем."
        };

        int tableTop = topHeight;
        int currentX = margin;
        g.setFont(bold);
        FontMetrics headerMetrics = g.getFontMetrics();
        for (int i = 0; i < headers.length; i++) {
            drawHeaderCell(g, headers[i], currentX, tableTop, columnWidths[i], headerHeight, headerMetrics);
            currentX += columnWidths[i];
        }

        g.setFont(regular);
        FontMetrics regularMetrics = g.getFontMetrics();
        int rowY = tableTop + headerHeight;
        int rowNo = 1;
        for (Map<String, Object> rowData : rows) {
            currentX = margin;
            String[] values = {
                    formatSmartNumber(rowNo++),
                    stringValue(rowData.get("project_name")),
                    stringValue(rowData.get("status_name")),
                    stringValue(rowData.get("customer_name")),
                    stringValue(rowData.get("address")),
                    formatIsoDate(rowData.get("start_date")),
                    formatIsoDate(rowData.get("end_date")),
                    formatSmartNumber(toDouble(rowData.get("planned_budget"))),
                    formatSmartNumber(toDouble(rowData.get("actual_budget"))),
                    formatSmartNumber(toDouble(rowData.get("budget_percent"))),
                    formatSmartNumber(toDouble(rowData.get("progress_percent"))),
                    formatSmartNumber(toDouble(rowData.get("blocks_count"))),
                    formatSmartNumber(toDouble(rowData.get("total_area"))),
                    formatSmartNumber(toDouble(rowData.get("sale_area"))),
                    formatSmartNumber(toDouble(rowData.get("material_requests_count"))),
                    formatSmartNumber(toDouble(rowData.get("purchase_orders_count"))),
                    formatSmartNumber(toDouble(rowData.get("signed_work_performed_count"))),
                    formatSmartNumber(
                            toDouble(rowData.get("material_write_off_count"))
                                    + toDouble(rowData.get("mbp_write_off_count"))
                                    + toDouble(rowData.get("processing_write_off_count"))
                    ),
                    formatSmartNumber(toDouble(rowData.get("transfer_count")))
            };

            for (int i = 0; i < values.length; i++) {
                g.drawRect(currentX, rowY, columnWidths[i], rowHeight);
                if (i >= 1 && i <= 4) {
                    drawLeftString(g, values[i], currentX + 4, rowY, columnWidths[i] - 8, rowHeight, regularMetrics);
                } else {
                    drawCenteredString(g, values[i], currentX, rowY, columnWidths[i], rowHeight, regularMetrics);
                }
                currentX += columnWidths[i];
            }
            rowY += rowHeight;
        }

        if (rows.isEmpty()) {
            currentX = margin;
            for (int columnWidth : columnWidths) {
                g.drawRect(currentX, rowY, columnWidth, rowHeight);
                currentX += columnWidth;
            }
            drawCenteredString(g, "Проектов для отображения нет", margin, rowY, tableWidth, rowHeight, regularMetrics);
        }

        g.dispose();
        return image;
    }

    private byte[] generateScheduleXlsx(ReportData reportData) throws Exception {
        ScheduleRenderModel model = buildScheduleModel(reportData);

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet(sanitizeSheetName(model.sheetName()));

            CellStyle baseStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle leftStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, null);
            CellStyle filledStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            applyHexFill(filledStyle, "#5B8FD1");

            int totalColumns = 2 + model.slots().size();
            for (int col = 0; col < totalColumns; col++) {
                sheet.setColumnWidth(col, col == 0 ? 1800 : col == 1 ? 11000 : 900);
            }

            int rowIndex = 0;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "Заказчик: " + model.customerName(), metaStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "Объект: " + model.projectName() + ", " + model.blockName(), metaStyle);
            rowIndex++;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, model.reportTitle(), titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "по объекту: " + model.projectName() + " " + model.blockName(), titleStyle);
            rowIndex++;

            Row yearRow = sheet.createRow(rowIndex++);
            Row monthRow = sheet.createRow(rowIndex++);
            Row slotRow = sheet.createRow(rowIndex++);

            writeCell(yearRow, 0, "№", headerStyle);
            writeCell(yearRow, 1, model.nameColumn(), headerStyle);
            writeCell(monthRow, 0, "", headerStyle);
            writeCell(monthRow, 1, "", headerStyle);
            writeCell(slotRow, 0, "", headerStyle);
            writeCell(slotRow, 1, "", headerStyle);

            mergeCellsIfNeeded(sheet, yearRow.getRowNum(), 0, slotRow.getRowNum(), 0);
            mergeCellsIfNeeded(sheet, yearRow.getRowNum(), 1, slotRow.getRowNum(), 1);

            int column = 2;
            int yearStart = column;
            String currentYear = model.slots().isEmpty() ? "" : model.slots().get(0).yearLabel();
            String currentMonth = model.slots().isEmpty() ? "" : model.slots().get(0).monthLabel();
            int monthStart = column;

            for (int i = 0; i < model.slots().size(); i++) {
                ScheduleSlot slot = model.slots().get(i);
                writeCell(slotRow, column, slot.slotLabel(), headerStyle);

                if (!slot.yearLabel().equals(currentYear)) {
                    mergeCellsIfNeeded(sheet, yearRow.getRowNum(), yearStart, column - 1);
                    writeCell(yearRow, yearStart, currentYear, headerStyle);
                    currentYear = slot.yearLabel();
                    yearStart = column;
                }

                if (!slot.monthLabel().equals(currentMonth)) {
                    mergeCellsIfNeeded(sheet, monthRow.getRowNum(), monthStart, column - 1);
                    writeCell(monthRow, monthStart, currentMonth, headerStyle);
                    currentMonth = slot.monthLabel();
                    monthStart = column;
                }

                column++;
            }

            if (!model.slots().isEmpty()) {
                mergeCellsIfNeeded(sheet, yearRow.getRowNum(), yearStart, column - 1);
                writeCell(yearRow, yearStart, currentYear, headerStyle);
                mergeCellsIfNeeded(sheet, monthRow.getRowNum(), monthStart, column - 1);
                writeCell(monthRow, monthStart, currentMonth, headerStyle);
            }

            for (ScheduleRowData rowData : model.rows()) {
                Row row = sheet.createRow(rowIndex++);
                writeCell(row, 0, String.valueOf(rowData.rowNo()), baseStyle);
                writeCell(row, 1, rowData.label(), leftStyle);
                for (int i = 0; i < model.slots().size(); i++) {
                    writeCell(row, i + 2, "", rowData.filled().get(i) ? filledStyle : baseStyle);
                }
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }

    private byte[] generateForm19XlsxDetailed(ReportData reportData) throws Exception {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Ф-19");

            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, null);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle integerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);

            titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            headerStyle.setWrapText(true);
            headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            textStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            integerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            decimalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            integerStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));
            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));

            int[] widths = {
                    1600, 11000, 2200,
                    2600, 2200,
                    2600, 2200,
                    2600, 2200,
                    3200, 2200,
                    3200, 2200,
                    3600, 2200,
                    2600, 2200
            };
            for (int i = 0; i < widths.length; i++) {
                sheet.setColumnWidth(i, widths[i]);
            }

            int lastColumn = 16;
            int rowIndex = 0;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "ОТЧЕТ Ф-19", titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "За " + stringValue(header.get("period_label")), titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "Заказчик: " + stringValue(header.get("customer_name")), metaStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "Объект: " + stringValue(header.get("project_name")) + ", " + stringValue(header.get("warehouse_name")), metaStyle);
            rowIndex++;

            Row topHeader = sheet.createRow(rowIndex++);
            Row middleHeader = sheet.createRow(rowIndex++);
            Row bottomHeader = sheet.createRow(rowIndex++);
            topHeader.setHeightInPoints(28);
            middleHeader.setHeightInPoints(28);
            bottomHeader.setHeightInPoints(22);

            for (int col = 0; col <= lastColumn; col++) {
                writeCell(topHeader, col, "", headerStyle);
                writeCell(middleHeader, col, "", headerStyle);
                writeCell(bottomHeader, col, "", headerStyle);
            }

            writeCell(topHeader, 0, "№", headerStyle);
            writeCell(topHeader, 1, "Наименование материалов", headerStyle);
            writeCell(topHeader, 2, "ЕИ", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 0, bottomHeader.getRowNum(), 0);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 1, bottomHeader.getRowNum(), 1);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 2, bottomHeader.getRowNum(), 2);

            writeCell(topHeader, 3, "Ост на нач мес", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 3, middleHeader.getRowNum(), 4);
            writeCell(topHeader, 5, "приход", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 5, middleHeader.getRowNum(), 6);
            writeCell(topHeader, 7, "Расход", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 7, 14);
            writeCell(topHeader, 13, "Ост на кон мес", headerStyle);
            mergeCellsIfNeeded(sheet, topHeader.getRowNum(), 15, middleHeader.getRowNum(), 16);
            writeCell(topHeader, 15, "Ост на кон мес", headerStyle);

            writeCell(middleHeader, 7, "Списание\nпо АВР", headerStyle);
            mergeCellsIfNeeded(sheet, middleHeader.getRowNum(), 7, 8);
            writeCell(middleHeader, 7, "Списание\nпо АВР", headerStyle);
            writeCell(middleHeader, 9, "Акт списания\nМБП", headerStyle);
            mergeCellsIfNeeded(sheet, middleHeader.getRowNum(), 9, 10);
            writeCell(middleHeader, 11, "Акт\nпереработки", headerStyle);
            mergeCellsIfNeeded(sheet, middleHeader.getRowNum(), 11, 12);
            writeCell(middleHeader, 13, "Перемещения\nпо складам", headerStyle);
            mergeCellsIfNeeded(sheet, middleHeader.getRowNum(), 13, 14);

            for (int col = 3; col <= lastColumn; col += 2) {
                writeCell(bottomHeader, col, "Кол", headerStyle);
                writeCell(bottomHeader, col + 1, "Цена", headerStyle);
            }

            int index = 1;
            for (Map<String, Object> rowData : rows) {
                Row row = sheet.createRow(rowIndex++);
                writeSmartNumericCell(row, 0, index++, integerStyle, decimalStyle);
                writeCell(row, 1, stringValue(rowData.get("material_name")), textStyle);
                writeCell(row, 2, stringValue(rowData.get("unit_name")), centerStyle);

                writeSmartNumericCell(row, 3, toDouble(rowData.get("opening_quantity")), integerStyle, decimalStyle);
                writeCell(row, 4, "", centerStyle);
                writeSmartNumericCell(row, 5, toDouble(rowData.get("incoming_quantity")), integerStyle, decimalStyle);
                writeCell(row, 6, "", centerStyle);
                writeSmartNumericCell(row, 7, toDouble(rowData.get("form29_quantity")), integerStyle, decimalStyle);
                writeCell(row, 8, "", centerStyle);
                writeSmartNumericCell(row, 9, toDouble(rowData.get("mbp_quantity")), integerStyle, decimalStyle);
                writeCell(row, 10, "", centerStyle);
                writeSmartNumericCell(row, 11, toDouble(rowData.get("processing_quantity")), integerStyle, decimalStyle);
                writeCell(row, 12, "", centerStyle);
                writeSmartNumericCell(row, 13, toDouble(rowData.get("transfer_out_quantity")), integerStyle, decimalStyle);
                writeCell(row, 14, "", centerStyle);
                writeSmartNumericCell(row, 15, toDouble(rowData.get("closing_quantity")), integerStyle, decimalStyle);
                writeCell(row, 16, "", centerStyle);
            }

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private byte[] generateForm19Xlsx(ReportData reportData) throws Exception {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Ф-19");

            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, null);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle integerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            short quantityFormat = workbook.createDataFormat().getFormat("#,##0.###");
            integerStyle.setDataFormat(quantityFormat);
            decimalStyle.setDataFormat(quantityFormat);

            sheet.setColumnWidth(0, 1600);
            sheet.setColumnWidth(1, 11000);
            sheet.setColumnWidth(2, 2200);
            sheet.setColumnWidth(3, 3600);
            sheet.setColumnWidth(4, 3600);
            sheet.setColumnWidth(5, 3600);
            sheet.setColumnWidth(6, 4200);
            sheet.setColumnWidth(7, 3600);

            int rowIndex = 0;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 7, "ОТЧЕТ Ф-19", titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 7, "За " + stringValue(header.get("period_label")), titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 7, "Заказчик: " + stringValue(header.get("customer_name")), metaStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 7, "Объект: " + stringValue(header.get("project_name")) + ", " + stringValue(header.get("warehouse_name")), metaStyle);
            rowIndex++;

            Row headerRow = sheet.createRow(rowIndex++);
            writeCell(headerRow, 0, "№", headerStyle);
            writeCell(headerRow, 1, "Наименование материалов", headerStyle);
            writeCell(headerRow, 2, "ЕИ", headerStyle);
            writeCell(headerRow, 3, "Ост на нач", headerStyle);
            writeCell(headerRow, 4, "Приход", headerStyle);
            writeCell(headerRow, 5, "Ф-29", headerStyle);
            writeCell(headerRow, 6, "Акт списания МБП", headerStyle);
            writeCell(headerRow, 7, "Ост на кон", headerStyle);

            for (Map<String, Object> rowData : rows) {
                Row row = sheet.createRow(rowIndex++);
                writeNumericCell(row, 0, toDouble(rowData.get("row_no")), centerStyle);
                writeCell(row, 1, stringValue(rowData.get("material_name")), textStyle);
                writeCell(row, 2, stringValue(rowData.get("unit_name")), centerStyle);
                writeSmartNumericCell(row, 3, toDouble(rowData.get("opening_quantity")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 4, toDouble(rowData.get("incoming_quantity")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 5, toDouble(rowData.get("form29_quantity")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 6, toDouble(rowData.get("mbp_quantity")), integerStyle, decimalStyle);
                writeSmartNumericCell(row, 7, toDouble(rowData.get("closing_quantity")), integerStyle, decimalStyle);
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }

    private BufferedImage renderScheduleImage(ReportData reportData) {
        ScheduleRenderModel model = buildScheduleModel(reportData);

        int margin = 24;
        int numberWidth = 42;
        int nameWidth = 310;
        int slotWidth = 22;
        int rowHeight = 24;
        int topHeight = 120;
        int headerHeight = 56;
        int footerHeight = 64;
        int tableWidth = numberWidth + nameWidth + model.slots().size() * slotWidth;
        int width = Math.max(520, margin * 2 + tableWidth);
        int height = topHeight + headerHeight + (Math.max(1, model.rows().size()) * rowHeight) + footerHeight + margin * 2;

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(Color.WHITE);
        g.fillRect(0, 0, width, height);

        java.awt.Font regular = new java.awt.Font("Arial", java.awt.Font.PLAIN, 12);
        java.awt.Font bold = new java.awt.Font("Arial", java.awt.Font.BOLD, 12);
        java.awt.Font title = new java.awt.Font("Arial", java.awt.Font.BOLD, 22);

        int x = margin;
        int y = margin + 10;

        g.setColor(Color.BLACK);
        g.setFont(bold);
        g.drawString("Заказчик: " + model.customerName(), x, y);
        y += 20;
        g.drawString("Объект: " + model.projectName() + ", " + model.blockName(), x, y);
        y += 36;

        FontMetrics titleMetrics = g.getFontMetrics(title);
        String reportTitle = model.reportTitle();
        g.setFont(title);
        g.drawString(reportTitle, (width - titleMetrics.stringWidth(reportTitle)) / 2, y);
        y += 26;

        g.setFont(bold);
        String subtitle = "по объекту: " + model.projectName() + " " + model.blockName();
        FontMetrics subtitleMetrics = g.getFontMetrics(bold);
        g.drawString(subtitle, (width - subtitleMetrics.stringWidth(subtitle)) / 2, y);
        y += 26;

        int tableTop = y;
        int tableLeft = x;
        int gridTop = tableTop + headerHeight;
        int gridLeft = tableLeft + numberWidth + nameWidth;

        g.setStroke(new BasicStroke(1f));
        g.setFont(bold);
        drawCenteredText(g, "№", tableLeft, tableTop, numberWidth, headerHeight);
        drawCenteredText(g, model.nameColumn(), tableLeft + numberWidth, tableTop, nameWidth, headerHeight);
        g.drawRect(tableLeft, tableTop, numberWidth, headerHeight);
        g.drawRect(tableLeft + numberWidth, tableTop, nameWidth, headerHeight);

        int slotX = gridLeft;
        String currentYear = null;
        int yearStartX = slotX;
        String currentMonth = null;
        int monthStartX = slotX;

        for (int i = 0; i < model.slots().size(); i++) {
            ScheduleSlot slot = model.slots().get(i);
            if (currentYear == null) {
                currentYear = slot.yearLabel();
                currentMonth = slot.monthLabel();
            }

            if (!slot.yearLabel().equals(currentYear)) {
                drawCenteredText(g, currentYear, yearStartX, tableTop, slotX - yearStartX, 18);
                currentYear = slot.yearLabel();
                yearStartX = slotX;
            }
            if (!slot.monthLabel().equals(currentMonth)) {
                drawCenteredText(g, currentMonth, monthStartX, tableTop + 18, slotX - monthStartX, 18);
                currentMonth = slot.monthLabel();
                monthStartX = slotX;
            }

            g.drawRect(slotX, tableTop, slotWidth, headerHeight);
            drawCenteredText(g, slot.slotLabel(), slotX, tableTop + 36, slotWidth, 20);
            slotX += slotWidth;
        }

        if (!model.slots().isEmpty()) {
            drawCenteredText(g, currentYear, yearStartX, tableTop, slotX - yearStartX, 18);
            drawCenteredText(g, currentMonth, monthStartX, tableTop + 18, slotX - monthStartX, 18);
        }

        int rowTop = gridTop;
        g.setFont(regular);
        for (ScheduleRowData rowData : model.rows()) {
            g.setColor(Color.BLACK);
            g.drawRect(tableLeft, rowTop, numberWidth, rowHeight);
            g.drawRect(tableLeft + numberWidth, rowTop, nameWidth, rowHeight);
            drawCenteredText(g, String.valueOf(rowData.rowNo()), tableLeft, rowTop, numberWidth, rowHeight);
            drawLeftText(g, rowData.label(), tableLeft + numberWidth + 6, rowTop + 16);

            int currentX = gridLeft;
            for (boolean filled : rowData.filled()) {
                if (filled) {
                    g.setColor(new Color(91, 143, 209));
                    g.fillRect(currentX + 1, rowTop + 1, slotWidth - 1, rowHeight - 1);
                }
                g.setColor(Color.BLACK);
                g.drawRect(currentX, rowTop, slotWidth, rowHeight);
                currentX += slotWidth;
            }
            rowTop += rowHeight;
        }

        g.setFont(bold);
        g.drawString("Главный инженер: ____________________", tableLeft, rowTop + 32);
        g.drawString("Инженер ПТО: ____________________", tableLeft, rowTop + 54);
        g.dispose();
        return image;
    }

    private BufferedImage renderForm19ImageDetailed(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        int margin = 24;
        int[] columnWidths = {40, 460, 70, 70, 60, 70, 60, 70, 60, 120, 60, 120, 60, 120, 60, 80, 60};
        int rowHeight = 26;
        int topHeight = 130;
        int headerTopHeight = 28;
        int headerMiddleHeight = 30;
        int headerBottomHeight = 22;
        int headerHeight = headerTopHeight + headerMiddleHeight + headerBottomHeight;
        int footerHeight = 32;
        int tableWidth = 0;
        for (int width : columnWidths) {
            tableWidth += width;
        }

        int width = Math.max(1320, margin * 2 + tableWidth);
        int height = topHeight + headerHeight + (Math.max(1, rows.size()) * rowHeight) + footerHeight + margin * 2;

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(Color.WHITE);
        g.fillRect(0, 0, width, height);

        java.awt.Font regular = new java.awt.Font("Arial", java.awt.Font.PLAIN, 12);
        java.awt.Font bold = new java.awt.Font("Arial", java.awt.Font.BOLD, 12);
        java.awt.Font title = new java.awt.Font("Arial", java.awt.Font.BOLD, 18);

        int x = margin;
        int y = margin + 10;

        g.setColor(Color.BLACK);
        g.setFont(title);
        String reportTitle = "ОТЧЕТ Ф-19";
        FontMetrics titleMetrics = g.getFontMetrics(title);
        g.drawString(reportTitle, (width - titleMetrics.stringWidth(reportTitle)) / 2, y);
        y += 24;

        g.setFont(bold);
        String subtitle = "За " + stringValue(header.get("period_label"));
        FontMetrics subtitleMetrics = g.getFontMetrics(bold);
        g.drawString(subtitle, (width - subtitleMetrics.stringWidth(subtitle)) / 2, y);
        y += 22;

        g.drawString("Заказчик: " + stringValue(header.get("customer_name")), x, y);
        y += 18;
        g.drawString("Объект: " + stringValue(header.get("project_name")) + ", " + stringValue(header.get("warehouse_name")), x, y);

        int tableX = margin;
        int tableY = topHeight;
        int[] colX = new int[columnWidths.length];
        colX[0] = tableX;
        for (int i = 1; i < columnWidths.length; i++) {
            colX[i] = colX[i - 1] + columnWidths[i - 1];
        }

        g.setFont(bold);
        FontMetrics headerMetrics = g.getFontMetrics(bold);

        drawHeaderCell(g, "№", colX[0], tableY, columnWidths[0], headerHeight, headerMetrics);
        drawHeaderCell(g, "Наименование материалов", colX[1], tableY, columnWidths[1], headerHeight, headerMetrics);
        drawHeaderCell(g, "ЕИ", colX[2], tableY, columnWidths[2], headerHeight, headerMetrics);

        drawHeaderCell(g, "Ост на нач мес", colX[3], tableY, columnWidths[3] + columnWidths[4], headerTopHeight + headerMiddleHeight, headerMetrics);
        drawHeaderCell(g, "приход", colX[5], tableY, columnWidths[5] + columnWidths[6], headerTopHeight + headerMiddleHeight, headerMetrics);
        drawHeaderCell(g, "Списание\nпо АВР", colX[7], tableY + headerTopHeight, columnWidths[7] + columnWidths[8], headerMiddleHeight, headerMetrics);
        drawHeaderCell(g, "Акт списания\nМБП", colX[9], tableY + headerTopHeight, columnWidths[9] + columnWidths[10], headerMiddleHeight, headerMetrics);

        drawHeaderCell(g, "Расход", colX[7], tableY, columnWidths[7] + columnWidths[8] + columnWidths[9] + columnWidths[10] + columnWidths[11] + columnWidths[12] + columnWidths[13] + columnWidths[14], headerTopHeight, headerMetrics);
        drawHeaderCell(g, "Акт\nпереработки", colX[11], tableY + headerTopHeight, columnWidths[11] + columnWidths[12], headerMiddleHeight, headerMetrics);
        drawHeaderCell(g, "Перемещения\nпо складам", colX[13], tableY + headerTopHeight, columnWidths[13] + columnWidths[14], headerMiddleHeight, headerMetrics);
        drawHeaderCell(g, "Ост на кон мес", colX[15], tableY, columnWidths[15] + columnWidths[16], headerTopHeight + headerMiddleHeight, headerMetrics);

        int bottomY = tableY + headerTopHeight + headerMiddleHeight;
        for (int col = 3; col < columnWidths.length; col += 2) {
            drawHeaderCell(g, "Кол", colX[col], bottomY, columnWidths[col], headerBottomHeight, headerMetrics);
            drawHeaderCell(g, "Цена", colX[col + 1], bottomY, columnWidths[col + 1], headerBottomHeight, headerMetrics);
        }

        int currentY = tableY + headerHeight;
        g.setFont(regular);
        FontMetrics regularMetrics = g.getFontMetrics(regular);

        for (Map<String, Object> row : rows) {
            String[] values = {
                    formatSmartNumber(toDouble(row.get("row_no"))),
                    stringValue(row.get("material_name")),
                    stringValue(row.get("unit_name")),
                    formatSmartNumber(toDouble(row.get("opening_quantity"))),
                    "",
                    formatSmartNumber(toDouble(row.get("incoming_quantity"))),
                    "",
                    formatSmartNumber(toDouble(row.get("form29_quantity"))),
                    "",
                    formatSmartNumber(toDouble(row.get("mbp_quantity"))),
                    "",
                    formatSmartNumber(toDouble(row.get("processing_quantity"))),
                    "",
                    formatSmartNumber(toDouble(row.get("transfer_out_quantity"))),
                    "",
                    formatSmartNumber(toDouble(row.get("closing_quantity"))),
                    ""
            };

            for (int i = 0; i < values.length; i++) {
                g.drawRect(colX[i], currentY, columnWidths[i], rowHeight);
                if (i == 1) {
                    drawLeftString(g, values[i], colX[i] + 4, currentY, columnWidths[i] - 8, rowHeight, regularMetrics);
                } else {
                    drawCenteredString(g, values[i], colX[i], currentY, columnWidths[i], rowHeight, regularMetrics);
                }
            }
            currentY += rowHeight;
        }

        g.dispose();
        return image;
    }

    private BufferedImage renderMbpWriteOffImage(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        int margin = 28;
        int[] columnWidths = {46, 360, 58, 72, 108, 132};
        int tableWidth = 0;
        for (int width : columnWidths) {
            tableWidth += width;
        }
        int rowHeight = 24;
        int headerHeight = 44;
        int topHeight = 230;
        int footerHeight = 90;
        int width = margin * 2 + tableWidth;
        int height = Math.max(620, topHeight + headerHeight + Math.max(1, rows.size()) * rowHeight + footerHeight);

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(Color.WHITE);
        g.fillRect(0, 0, width, height);

        java.awt.Font regular = new java.awt.Font("Arial", java.awt.Font.PLAIN, 13);
        java.awt.Font bold = new java.awt.Font("Arial", java.awt.Font.BOLD, 13);
        java.awt.Font title = new java.awt.Font("Arial", java.awt.Font.BOLD, 16);
        java.awt.Font smallBold = new java.awt.Font("Arial", java.awt.Font.BOLD, 12);

        g.setColor(Color.BLACK);
        g.setFont(bold);
        FontMetrics boldMetrics = g.getFontMetrics();

        int approvalX = margin + 420;
        int y = 24;
        drawCenteredString(g, "\"УТВЕРЖДАЮ\"", approvalX, y, 260, 20, boldMetrics);
        y += 22;
        drawCenteredString(g, "Генеральный директор", approvalX, y, 260, 20, boldMetrics);
        y += 22;
        drawCenteredString(g, "ОсОО \"СК Дрим Хаус\"", approvalX, y, 260, 20, boldMetrics);
        y += 22;
        drawCenteredString(g, "Темиркулов Ш. Т.", approvalX, y, 260, 20, boldMetrics);
        y += 24;
        g.drawLine(approvalX + 40, y, approvalX + 220, y);

        y = 132;
        g.setFont(title);
        FontMetrics titleMetrics = g.getFontMetrics();
        drawCenteredString(g, "АКТ", margin, y, tableWidth, 22, titleMetrics);
        y += 26;
        g.setFont(bold);
        boldMetrics = g.getFontMetrics();
        drawCenteredString(g, "на списание малоценных и быстроизнашиваемых материалов", margin, y, tableWidth, 20, boldMetrics);
        y += 24;
        drawCenteredString(g, "за " + stringValue(header.get("period_label")), margin, y, tableWidth, 20, boldMetrics);
        y += 30;
        drawLeftString(g, "Объект: " + stringValue(header.get("project_name")) + projectAddressSuffix(header), margin, y, tableWidth, 18, boldMetrics);

        int tableTop = topHeight;
        int currentX = margin;
        String[] headers = {"№\nп/п", "Наименование материала", "Ед.\nизм.", "Кол-\nво", "Дата\nсписания", "Примечание"};
        g.setFont(smallBold);
        FontMetrics headerMetrics = g.getFontMetrics();
        for (int i = 0; i < headers.length; i++) {
            drawHeaderCell(g, headers[i], currentX, tableTop, columnWidths[i], headerHeight, headerMetrics);
            currentX += columnWidths[i];
        }

        g.setFont(regular);
        FontMetrics regularMetrics = g.getFontMetrics();
        int rowY = tableTop + headerHeight;
        for (Map<String, Object> rowData : rows) {
            currentX = margin;
            String[] values = {
                    formatSmartNumber(toDouble(rowData.get("row_no"))),
                    stringValue(rowData.get("material_name")),
                    stringValue(rowData.get("unit_name")),
                    formatSmartNumber(toDouble(rowData.get("quantity"))),
                    stringValue(rowData.get("posted_at_label")),
                    stringValue(rowData.get("note"))
            };
            for (int i = 0; i < values.length; i++) {
                g.drawRect(currentX, rowY, columnWidths[i], rowHeight);
                if (i == 1 || i == 5) {
                    drawLeftString(g, values[i], currentX + 4, rowY, columnWidths[i] - 8, rowHeight, regularMetrics);
                } else {
                    drawCenteredString(g, values[i], currentX, rowY, columnWidths[i], rowHeight, regularMetrics);
                }
                currentX += columnWidths[i];
            }
            rowY += rowHeight;
        }

        if (rows.isEmpty()) {
            currentX = margin;
            for (int columnWidth : columnWidths) {
                g.drawRect(currentX, rowY, columnWidth, rowHeight);
                currentX += columnWidth;
            }
            drawCenteredString(g, "Списаний МБП за период нет", margin, rowY, tableWidth, rowHeight, regularMetrics);
            rowY += rowHeight;
        }

        rowY += 34;
        g.setFont(bold);
        boldMetrics = g.getFontMetrics();
        drawLeftString(g, "Прораб : __________________________", margin + 40, rowY, 360, 24, boldMetrics);

        g.dispose();
        return image;
    }

    private void drawHeaderCell(Graphics2D g, String text, int x, int y, int width, int height, FontMetrics metrics) {
        g.drawRect(x, y, width, height);
        String[] lines = (text == null ? "" : text).split("\\n", -1);
        int totalTextHeight = lines.length * metrics.getHeight();
        int baseline = y + Math.max(0, (height - totalTextHeight) / 2) + metrics.getAscent();
        for (String line : lines) {
            int textX = x + Math.max(0, (width - metrics.stringWidth(line)) / 2);
            g.drawString(line, textX, baseline);
            baseline += metrics.getHeight();
        }
    }

    private BufferedImage renderForm19Image(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        int margin = 24;
        int[] columnWidths = {40, 460, 70, 100, 100, 100, 120, 100};
        int rowHeight = 28;
        int topHeight = 120;
        int headerHeight = 34;
        int footerHeight = 32;
        int tableWidth = 0;
        for (int width : columnWidths) {
            tableWidth += width;
        }

        int width = Math.max(520, margin * 2 + tableWidth);
        int height = topHeight + headerHeight + (Math.max(1, rows.size()) * rowHeight) + footerHeight + margin * 2;

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(Color.WHITE);
        g.fillRect(0, 0, width, height);

        java.awt.Font regular = new java.awt.Font("Arial", java.awt.Font.PLAIN, 12);
        java.awt.Font bold = new java.awt.Font("Arial", java.awt.Font.BOLD, 12);
        java.awt.Font title = new java.awt.Font("Arial", java.awt.Font.BOLD, 18);

        int x = margin;
        int y = margin + 10;

        g.setColor(Color.BLACK);
        g.setFont(title);
        String reportTitle = "ОТЧЕТ Ф-19";
        FontMetrics titleMetrics = g.getFontMetrics(title);
        g.drawString(reportTitle, (width - titleMetrics.stringWidth(reportTitle)) / 2, y);
        y += 24;

        g.setFont(bold);
        String subtitle = "За " + stringValue(header.get("period_label"));
        FontMetrics subtitleMetrics = g.getFontMetrics(bold);
        g.drawString(subtitle, (width - subtitleMetrics.stringWidth(subtitle)) / 2, y);
        y += 22;

        g.drawString("Заказчик: " + stringValue(header.get("customer_name")), x, y);
        y += 18;
        g.drawString("Объект: " + stringValue(header.get("project_name")) + ", " + stringValue(header.get("warehouse_name")), x, y);

        int tableX = margin;
        int tableY = topHeight;

        String[] headers = {
                "№", "Наименование материалов", "ЕИ", "Ост на нач",
                "Приход", "Ф-29", "Акт списания МБП", "Ост на кон"
        };

        int currentX = tableX;
        g.setFont(bold);
        FontMetrics headerMetrics = g.getFontMetrics(bold);
        for (int i = 0; i < headers.length; i++) {
            g.drawRect(currentX, tableY, columnWidths[i], headerHeight);
            drawCenteredString(g, headers[i], currentX, tableY, columnWidths[i], headerHeight, headerMetrics);
            currentX += columnWidths[i];
        }

        int currentY = tableY + headerHeight;
        g.setFont(regular);
        FontMetrics regularMetrics = g.getFontMetrics(regular);

        for (Map<String, Object> row : rows) {
            currentX = tableX;
            String[] values = {
                    formatSmartNumber(toDouble(row.get("row_no"))),
                    stringValue(row.get("material_name")),
                    stringValue(row.get("unit_name")),
                    formatSmartNumber(toDouble(row.get("opening_quantity"))),
                    formatSmartNumber(toDouble(row.get("incoming_quantity"))),
                    formatSmartNumber(toDouble(row.get("form29_quantity"))),
                    formatSmartNumber(toDouble(row.get("mbp_quantity"))),
                    formatSmartNumber(toDouble(row.get("closing_quantity")))
            };

            for (int i = 0; i < values.length; i++) {
                g.drawRect(currentX, currentY, columnWidths[i], rowHeight);
                if (i == 1) {
                    drawLeftString(g, values[i], currentX + 4, currentY, columnWidths[i] - 8, rowHeight, regularMetrics);
                } else {
                    drawCenteredString(g, values[i], currentX, currentY, columnWidths[i], rowHeight, regularMetrics);
                }
                currentX += columnWidths[i];
            }
            currentY += rowHeight;
        }

        g.dispose();
        return image;
    }

    private ScheduleRenderModel buildScheduleModel(ReportData reportData) {
        Map header = reportData.header();
        Map data = reportData.data();

        LocalDate dateFrom = parseLocalDate(header.get("date_from"));
        LocalDate dateTo = parseLocalDate(header.get("date_to"));
        if (dateFrom == null || dateTo == null) {
            throw new IllegalStateException("Schedule header does not contain a valid date range");
        }

        List<ScheduleSlot> slots = buildScheduleSlots(dateFrom, dateTo);
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());
        List<ScheduleRowData> renderedRows = new ArrayList<>();

        int rowNo = 1;
        for (Map<String, Object> row : rows) {
            LocalDate stageStart = parseLocalDate(row.get("start_date"));
            LocalDate stageEnd = parseLocalDate(row.get("end_date"));
            if (stageStart == null && stageEnd != null) {
                stageStart = stageEnd;
            }
            if (stageEnd == null && stageStart != null) {
                stageEnd = stageStart;
            }

            List<Boolean> filled = new ArrayList<>();
            for (ScheduleSlot slot : slots) {
                boolean overlap = stageStart != null
                        && stageEnd != null
                        && !slot.end().isBefore(stageStart)
                        && !slot.start().isAfter(stageEnd);
                filled.add(overlap);
            }

            renderedRows.add(new ScheduleRowData(rowNo++, stringValue(row.get("stage_name")), filled));
        }

        return new ScheduleRenderModel(
                stringValue(header.get("project_name")),
                stringValue(header.get("block_name")),
                stringValue(header.get("customer_name")),
                scheduleReportTitle(header),
                scheduleNameColumn(header),
                scheduleSheetName(header),
                slots,
                renderedRows
        );
    }

    private byte[] generateProjectScheduleXlsx(ReportData reportData) throws Exception {
        ScheduleRenderModel model = buildProjectScheduleModel(reportData);
        List<ScheduleRenderModel> blockModels = buildProjectBlockScheduleModels(reportData);

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("График");

            CellStyle baseStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle leftStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, null);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, null);
            CellStyle filledStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle sectionStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            applyHexFill(filledStyle, "#5B8FD1");
            applyHexFill(sectionStyle, "#F4F4F5");

            int totalColumns = 2 + model.slots().size();
            for (int col = 0; col < totalColumns; col++) {
                sheet.setColumnWidth(col, col == 0 ? 1800 : col == 1 ? 11000 : 900);
            }

            int rowIndex = 0;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "Заказчик: " + model.customerName(), metaStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "Объект: " + model.projectName(), metaStyle);
            rowIndex++;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "КАЛЕНДАРНЫЙ ГРАФИК ПРОИЗВОДСТВА РАБОТ", titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "по объекту: " + model.projectName(), titleStyle);
            rowIndex++;

            Row yearRow = sheet.createRow(rowIndex++);
            Row monthRow = sheet.createRow(rowIndex++);
            Row slotRow = sheet.createRow(rowIndex++);

            writeCell(yearRow, 0, "№", headerStyle);
            writeCell(yearRow, 1, "Наименование работ", headerStyle);
            writeCell(monthRow, 0, "", headerStyle);
            writeCell(monthRow, 1, "", headerStyle);
            writeCell(slotRow, 0, "", headerStyle);
            writeCell(slotRow, 1, "", headerStyle);

            mergeCellsIfNeeded(sheet, yearRow.getRowNum(), 0, slotRow.getRowNum(), 0);
            mergeCellsIfNeeded(sheet, yearRow.getRowNum(), 1, slotRow.getRowNum(), 1);

            int column = 2;
            int yearStart = column;
            String currentYear = model.slots().isEmpty() ? "" : model.slots().get(0).yearLabel();
            String currentMonth = model.slots().isEmpty() ? "" : model.slots().get(0).monthLabel();
            int monthStart = column;

            for (ScheduleSlot slot : model.slots()) {
                writeCell(slotRow, column, slot.slotLabel(), headerStyle);

                if (!slot.yearLabel().equals(currentYear)) {
                    mergeCellsIfNeeded(sheet, yearRow.getRowNum(), yearStart, column - 1);
                    writeCell(yearRow, yearStart, currentYear, headerStyle);
                    currentYear = slot.yearLabel();
                    yearStart = column;
                }

                if (!slot.monthLabel().equals(currentMonth)) {
                    mergeCellsIfNeeded(sheet, monthRow.getRowNum(), monthStart, column - 1);
                    writeCell(monthRow, monthStart, currentMonth, headerStyle);
                    currentMonth = slot.monthLabel();
                    monthStart = column;
                }

                column++;
            }

            if (!model.slots().isEmpty()) {
                mergeCellsIfNeeded(sheet, yearRow.getRowNum(), yearStart, column - 1);
                writeCell(yearRow, yearStart, currentYear, headerStyle);
                mergeCellsIfNeeded(sheet, monthRow.getRowNum(), monthStart, column - 1);
                writeCell(monthRow, monthStart, currentMonth, headerStyle);
            }

            for (ScheduleRowData rowData : model.rows()) {
                Row row = sheet.createRow(rowIndex++);
                if (rowData.rowNo() == 0) {
                    writeCell(row, 0, "", sectionStyle);
                    writeCell(row, 1, rowData.label(), sectionStyle);
                    for (int i = 0; i < model.slots().size(); i++) {
                        writeCell(row, i + 2, "", sectionStyle);
                    }
                } else {
                    writeCell(row, 0, String.valueOf(rowData.rowNo()), baseStyle);
                    writeCell(row, 1, rowData.label(), leftStyle);
                    for (int i = 0; i < model.slots().size(); i++) {
                        writeCell(row, i + 2, "", rowData.filled().get(i) ? filledStyle : baseStyle);
                    }
                }
            }

            for (ScheduleRenderModel blockModel : blockModels) {
                writeScheduleSheet(
                        workbook.createSheet(sanitizeSheetName(blockModel.blockName())),
                        blockModel,
                        baseStyle,
                        leftStyle,
                        headerStyle,
                        titleStyle,
                        metaStyle,
                        filledStyle,
                        sectionStyle
                );
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }

    private void writeScheduleSheet(
            Sheet sheet,
            ScheduleRenderModel model,
            CellStyle baseStyle,
            CellStyle leftStyle,
            CellStyle headerStyle,
            CellStyle titleStyle,
            CellStyle metaStyle,
            CellStyle filledStyle,
            CellStyle sectionStyle
    ) {
        int totalColumns = 2 + model.slots().size();
        for (int col = 0; col < totalColumns; col++) {
            sheet.setColumnWidth(col, col == 0 ? 1800 : col == 1 ? 11000 : 900);
        }

        int rowIndex = 0;
        rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "Заказчик: " + model.customerName(), metaStyle);
        rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "Объект: " + model.projectName() + (model.blockName().isBlank() ? "" : ", " + model.blockName()), metaStyle);
        rowIndex++;
        rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, model.reportTitle(), titleStyle);
        rowIndex = writeMergedValue(sheet, rowIndex, 0, totalColumns - 1, "по объекту: " + model.projectName() + (model.blockName().isBlank() ? "" : " " + model.blockName()), titleStyle);
        rowIndex++;

        Row yearRow = sheet.createRow(rowIndex++);
        Row monthRow = sheet.createRow(rowIndex++);
        Row slotRow = sheet.createRow(rowIndex++);

        writeCell(yearRow, 0, "№", headerStyle);
        writeCell(yearRow, 1, model.nameColumn(), headerStyle);
        writeCell(monthRow, 0, "", headerStyle);
        writeCell(monthRow, 1, "", headerStyle);
        writeCell(slotRow, 0, "", headerStyle);
        writeCell(slotRow, 1, "", headerStyle);

        mergeCellsIfNeeded(sheet, yearRow.getRowNum(), 0, slotRow.getRowNum(), 0);
        mergeCellsIfNeeded(sheet, yearRow.getRowNum(), 1, slotRow.getRowNum(), 1);

        int column = 2;
        int yearStart = column;
        String currentYear = model.slots().isEmpty() ? "" : model.slots().get(0).yearLabel();
        String currentMonth = model.slots().isEmpty() ? "" : model.slots().get(0).monthLabel();
        int monthStart = column;

        for (ScheduleSlot slot : model.slots()) {
            writeCell(slotRow, column, slot.slotLabel(), headerStyle);

            if (!slot.yearLabel().equals(currentYear)) {
                mergeCellsIfNeeded(sheet, yearRow.getRowNum(), yearStart, column - 1);
                writeCell(yearRow, yearStart, currentYear, headerStyle);
                currentYear = slot.yearLabel();
                yearStart = column;
            }

            if (!slot.monthLabel().equals(currentMonth)) {
                mergeCellsIfNeeded(sheet, monthRow.getRowNum(), monthStart, column - 1);
                writeCell(monthRow, monthStart, currentMonth, headerStyle);
                currentMonth = slot.monthLabel();
                monthStart = column;
            }

            column++;
        }

        if (!model.slots().isEmpty()) {
            mergeCellsIfNeeded(sheet, yearRow.getRowNum(), yearStart, column - 1);
            writeCell(yearRow, yearStart, currentYear, headerStyle);
            mergeCellsIfNeeded(sheet, monthRow.getRowNum(), monthStart, column - 1);
            writeCell(monthRow, monthStart, currentMonth, headerStyle);
        }

        for (ScheduleRowData rowData : model.rows()) {
            Row row = sheet.createRow(rowIndex++);
            if (rowData.rowNo() == 0) {
                writeCell(row, 0, "", sectionStyle);
                writeCell(row, 1, rowData.label(), sectionStyle);
                for (int i = 0; i < model.slots().size(); i++) {
                    writeCell(row, i + 2, "", sectionStyle);
                }
            } else {
                writeCell(row, 0, String.valueOf(rowData.rowNo()), baseStyle);
                writeCell(row, 1, rowData.label(), leftStyle);
                for (int i = 0; i < model.slots().size(); i++) {
                    writeCell(row, i + 2, "", rowData.filled().get(i) ? filledStyle : baseStyle);
                }
            }
        }
    }

    private List<ScheduleRenderModel> buildProjectBlockScheduleModels(ReportData reportData) {
        Map header = reportData.header();
        Map data = reportData.data();

        LocalDate dateFrom = parseLocalDate(header.get("date_from"));
        LocalDate dateTo = parseLocalDate(header.get("date_to"));
        if (dateFrom == null || dateTo == null) {
            throw new IllegalStateException("Schedule header does not contain a valid date range");
        }

        List<ScheduleSlot> slots = buildScheduleSlots(dateFrom, dateTo);
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());
        Map<String, List<Map<String, Object>>> groupedRows = new LinkedHashMap<>();
        Map<String, String> blockNames = new LinkedHashMap<>();

        for (Map<String, Object> row : rows) {
            String blockId = stringValue(row.get("block_id"));
            groupedRows.computeIfAbsent(blockId, key -> new ArrayList<>()).add(row);
            blockNames.putIfAbsent(blockId, stringValue(row.get("block_name")));
        }

        List<ScheduleRenderModel> result = new ArrayList<>();

        for (Map.Entry<String, List<Map<String, Object>>> entry : groupedRows.entrySet()) {
            List<ScheduleRowData> renderedRows = new ArrayList<>();
            int rowNo = 1;

            for (Map<String, Object> row : entry.getValue()) {
                LocalDate stageStart = parseLocalDate(row.get("start_date"));
                LocalDate stageEnd = parseLocalDate(row.get("end_date"));
                if (stageStart == null && stageEnd != null) {
                    stageStart = stageEnd;
                }
                if (stageEnd == null && stageStart != null) {
                    stageEnd = stageStart;
                }

                List<Boolean> filled = new ArrayList<>();
                for (ScheduleSlot slot : slots) {
                    boolean overlap = stageStart != null
                            && stageEnd != null
                            && !slot.end().isBefore(stageStart)
                            && !slot.start().isAfter(stageEnd);
                    filled.add(overlap);
                }

                renderedRows.add(new ScheduleRowData(rowNo++, stringValue(row.get("stage_name")), filled));
            }

            result.add(new ScheduleRenderModel(
                    stringValue(header.get("project_name")),
                    blockNames.getOrDefault(entry.getKey(), ""),
                    stringValue(header.get("customer_name")),
                    scheduleReportTitle(header),
                    scheduleNameColumn(header),
                    scheduleSheetName(header),
                    slots,
                    renderedRows
            ));
        }

        return result;
    }

    private BufferedImage renderProjectScheduleImage(ReportData reportData) {
        ScheduleRenderModel model = buildProjectScheduleModel(reportData);

        int margin = 24;
        int numberWidth = 42;
        int nameWidth = 310;
        int slotWidth = 22;
        int rowHeight = 24;
        int topHeight = 120;
        int headerHeight = 56;
        int footerHeight = 64;
        int tableWidth = numberWidth + nameWidth + model.slots().size() * slotWidth;
        int width = Math.max(520, margin * 2 + tableWidth);
        int height = topHeight + headerHeight + (Math.max(1, model.rows().size()) * rowHeight) + footerHeight + margin * 2;

        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(Color.WHITE);
        g.fillRect(0, 0, width, height);

        java.awt.Font regular = new java.awt.Font("Arial", java.awt.Font.PLAIN, 12);
        java.awt.Font bold = new java.awt.Font("Arial", java.awt.Font.BOLD, 12);
        java.awt.Font title = new java.awt.Font("Arial", java.awt.Font.BOLD, 22);

        int x = margin;
        int y = margin + 10;

        g.setColor(Color.BLACK);
        g.setFont(bold);
        g.drawString("Заказчик: " + model.customerName(), x, y);
        y += 20;
        g.drawString("Объект: " + model.projectName(), x, y);
        y += 36;

        FontMetrics titleMetrics = g.getFontMetrics(title);
        String reportTitle = "КАЛЕНДАРНЫЙ ГРАФИК ПРОИЗВОДСТВА РАБОТ";
        g.setFont(title);
        g.drawString(reportTitle, (width - titleMetrics.stringWidth(reportTitle)) / 2, y);
        y += 26;

        g.setFont(bold);
        String subtitle = "по объекту: " + model.projectName();
        FontMetrics subtitleMetrics = g.getFontMetrics(bold);
        g.drawString(subtitle, (width - subtitleMetrics.stringWidth(subtitle)) / 2, y);
        y += 26;

        int tableTop = y;
        int tableLeft = x;
        int gridTop = tableTop + headerHeight;
        int gridLeft = tableLeft + numberWidth + nameWidth;

        g.setStroke(new BasicStroke(1f));
        g.setFont(bold);
        drawCenteredText(g, "№", tableLeft, tableTop, numberWidth, headerHeight);
        drawCenteredText(g, "Наименование работ", tableLeft + numberWidth, tableTop, nameWidth, headerHeight);
        g.drawRect(tableLeft, tableTop, numberWidth, headerHeight);
        g.drawRect(tableLeft + numberWidth, tableTop, nameWidth, headerHeight);

        int slotX = gridLeft;
        String currentYear = null;
        int yearStartX = slotX;
        String currentMonth = null;
        int monthStartX = slotX;

        for (ScheduleSlot slot : model.slots()) {
            if (currentYear == null) {
                currentYear = slot.yearLabel();
                currentMonth = slot.monthLabel();
            }

            if (!slot.yearLabel().equals(currentYear)) {
                drawCenteredText(g, currentYear, yearStartX, tableTop, slotX - yearStartX, 18);
                currentYear = slot.yearLabel();
                yearStartX = slotX;
            }
            if (!slot.monthLabel().equals(currentMonth)) {
                drawCenteredText(g, currentMonth, monthStartX, tableTop + 18, slotX - monthStartX, 18);
                currentMonth = slot.monthLabel();
                monthStartX = slotX;
            }

            g.drawRect(slotX, tableTop, slotWidth, headerHeight);
            drawCenteredText(g, slot.slotLabel(), slotX, tableTop + 36, slotWidth, 20);
            slotX += slotWidth;
        }

        if (!model.slots().isEmpty()) {
            drawCenteredText(g, currentYear, yearStartX, tableTop, slotX - yearStartX, 18);
            drawCenteredText(g, currentMonth, monthStartX, tableTop + 18, slotX - monthStartX, 18);
        }

        int rowTop = gridTop;
        g.setFont(regular);
        for (ScheduleRowData rowData : model.rows()) {
            if (rowData.rowNo() == 0) {
                g.setColor(new Color(244, 244, 245));
                g.fillRect(tableLeft, rowTop, tableWidth, rowHeight);
                g.setColor(Color.BLACK);
                g.drawRect(tableLeft, rowTop, numberWidth, rowHeight);
                g.drawRect(tableLeft + numberWidth, rowTop, nameWidth, rowHeight);
                int currentX = gridLeft;
                for (int i = 0; i < model.slots().size(); i++) {
                    g.drawRect(currentX, rowTop, slotWidth, rowHeight);
                    currentX += slotWidth;
                }
                drawLeftText(g, rowData.label(), tableLeft + numberWidth + 6, rowTop + 16);
            } else {
                g.setColor(Color.BLACK);
                g.drawRect(tableLeft, rowTop, numberWidth, rowHeight);
                g.drawRect(tableLeft + numberWidth, rowTop, nameWidth, rowHeight);
                drawCenteredText(g, String.valueOf(rowData.rowNo()), tableLeft, rowTop, numberWidth, rowHeight);
                drawLeftText(g, rowData.label(), tableLeft + numberWidth + 6, rowTop + 16);

                int currentX = gridLeft;
                for (boolean filled : rowData.filled()) {
                    if (filled) {
                        g.setColor(new Color(91, 143, 209));
                        g.fillRect(currentX + 1, rowTop + 1, slotWidth - 1, rowHeight - 1);
                    }
                    g.setColor(Color.BLACK);
                    g.drawRect(currentX, rowTop, slotWidth, rowHeight);
                    currentX += slotWidth;
                }
            }
            rowTop += rowHeight;
        }

        g.setFont(bold);
        g.drawString("Главный инженер: ____________________", tableLeft, rowTop + 32);
        g.drawString("Инженер ПТО: ____________________", tableLeft, rowTop + 54);
        g.dispose();
        return image;
    }

    private ScheduleRenderModel buildProjectScheduleModel(ReportData reportData) {
        Map header = reportData.header();
        Map data = reportData.data();

        LocalDate dateFrom = parseLocalDate(header.get("date_from"));
        LocalDate dateTo = parseLocalDate(header.get("date_to"));
        if (dateFrom == null || dateTo == null) {
            throw new IllegalStateException("Schedule header does not contain a valid date range");
        }

        List<ScheduleSlot> slots = buildScheduleSlots(dateFrom, dateTo);
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());
        List<ScheduleRowData> renderedRows = new ArrayList<>();

        int rowNo = 1;
        String currentBlockId = null;
        for (Map<String, Object> row : rows) {
            String blockId = stringValue(row.get("block_id"));
            if (!blockId.equals(currentBlockId)) {
                currentBlockId = blockId;
                renderedRows.add(new ScheduleRowData(0, stringValue(row.get("block_name")), List.of()));
            }

            LocalDate stageStart = parseLocalDate(row.get("start_date"));
            LocalDate stageEnd = parseLocalDate(row.get("end_date"));
            if (stageStart == null && stageEnd != null) {
                stageStart = stageEnd;
            }
            if (stageEnd == null && stageStart != null) {
                stageEnd = stageStart;
            }

            List<Boolean> filled = new ArrayList<>();
            for (ScheduleSlot slot : slots) {
                boolean overlap = stageStart != null
                        && stageEnd != null
                        && !slot.end().isBefore(stageStart)
                        && !slot.start().isAfter(stageEnd);
                filled.add(overlap);
            }

            renderedRows.add(new ScheduleRowData(rowNo++, stringValue(row.get("stage_name")), filled));
        }

        return new ScheduleRenderModel(
                stringValue(header.get("project_name")),
                "",
                stringValue(header.get("customer_name")),
                scheduleReportTitle(header),
                scheduleNameColumn(header),
                scheduleSheetName(header),
                slots,
                renderedRows
        );
    }

    private List<ScheduleSlot> buildScheduleSlots(LocalDate dateFrom, LocalDate dateTo) {
        List<ScheduleSlot> slots = new ArrayList<>();
        YearMonth current = YearMonth.from(dateFrom.withDayOfMonth(1));
        YearMonth end = YearMonth.from(dateTo.withDayOfMonth(1));

        while (!current.isAfter(end)) {
            LocalDate monthStart = current.atDay(1);
            LocalDate firstEnd = current.atDay(Math.min(10, current.lengthOfMonth()));
            LocalDate secondStart = current.atDay(Math.min(11, current.lengthOfMonth()));
            LocalDate secondEnd = current.atDay(Math.min(20, current.lengthOfMonth()));
            LocalDate thirdStart = current.atDay(Math.min(21, current.lengthOfMonth()));
            LocalDate thirdEnd = current.atEndOfMonth();

            slots.add(new ScheduleSlot(monthStart, firstEnd, String.valueOf(current.getYear()), monthLabel(current), "10"));
            slots.add(new ScheduleSlot(secondStart, secondEnd, String.valueOf(current.getYear()), monthLabel(current), "10"));
            slots.add(new ScheduleSlot(thirdStart, thirdEnd, String.valueOf(current.getYear()), monthLabel(current), "10"));

            current = current.plusMonths(1);
        }

        return slots;
    }

    private String monthLabel(YearMonth yearMonth) {
        return yearMonth.getMonth().getDisplayName(java.time.format.TextStyle.SHORT_STANDALONE, new Locale("ru"))
                .replace(".", "");
    }

    private LocalDate parseLocalDate(Object value) {
        String text = stringValue(value);
        if (text.isBlank()) {
            return null;
        }
        if (text.length() >= 10) {
            text = text.substring(0, 10);
        }
        return LocalDate.parse(text, DateTimeFormatter.ISO_LOCAL_DATE);
    }

    private String formatIsoDate(Object value) {
        String text = stringValue(value);
        if (text.isBlank()) {
            return "";
        }
        if (text.length() >= 10) {
            return text.substring(0, 10);
        }
        return text;
    }

    private void drawCenteredText(Graphics2D g, String text, int x, int y, int width, int height) {
        FontMetrics metrics = g.getFontMetrics();
        int textX = x + Math.max(0, (width - metrics.stringWidth(text)) / 2);
        int textY = y + ((height - metrics.getHeight()) / 2) + metrics.getAscent();
        g.drawString(text, textX, textY);
    }

    private void drawLeftText(Graphics2D g, String text, int x, int baselineY) {
        g.drawString(text, x, baselineY);
    }

    public byte[] generateEstimateStage(ReportData reportData, String format) throws Exception {
        if ("html".equalsIgnoreCase(format)) {
            return generateEstimateStageHtml(reportData);
        }

        if (!"xlsx".equalsIgnoreCase(format)) {
            throw new IllegalArgumentException("Unsupported estimate stage format: " + format);
        }

        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> stages = (List<Map<String, Object>>) data.getOrDefault("stages", List.of());
        Map<String, Object> summary = (Map<String, Object>) data.getOrDefault("summary", Map.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, BorderStyle.NONE);
            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.NONE);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle moneyStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle totalStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);

            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            moneyStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            totalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));

            for (Map<String, Object> stage : stages) {
                writeEstimateStageSheet(workbook, stage, header, metaStyle, titleStyle, headerStyle, textStyle, centerStyle, decimalStyle, moneyStyle, totalStyle);
            }

            writeEstimateSummarySheet(workbook, summary, header, metaStyle, titleStyle, headerStyle, textStyle, centerStyle, moneyStyle, totalStyle);

            workbook.write(out);
            return out.toByteArray();
        }
    }

    private byte[] generateEstimateStageHtml(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> stages = (List<Map<String, Object>>) data.getOrDefault("stages", List.of());
        Map<String, Object> summary = (Map<String, Object>) data.getOrDefault("summary", Map.of());

        StringBuilder html = new StringBuilder();
        html.append("<!doctype html><html><head><meta charset=\"UTF-8\">");
        html.append("<style>");
        html.append("html,body{margin:0;padding:0;background:#f3f4f6;color:#111827;font-family:Arial,sans-serif;}");
        html.append("body{padding:16px;}");
        html.append(".page{width:max-content;min-width:100%;background:#fff;padding:20px;box-shadow:0 2px 10px rgba(0,0,0,.08);}");
        html.append(".meta{font-size:13px;font-weight:700;margin-bottom:4px;}");
        html.append(".title{font-size:22px;font-weight:700;text-align:center;margin:18px 0 14px;}");
        html.append(".stage-title{font-size:18px;font-weight:700;text-align:center;margin:20px 0 10px;}");
        html.append(".section-title{font-size:14px;font-weight:700;margin:10px 0 6px;}");
        html.append("table{border-collapse:collapse;width:max-content;min-width:100%;margin-bottom:16px;}");
        html.append("th,td{border:1px solid #000;padding:6px 8px;font-size:12px;vertical-align:middle;text-align:center;}");
        html.append("th{font-weight:700;background:#f8fafc;}");
        html.append(".left{text-align:left;}");
        html.append(".total td{font-weight:700;background:#fff7bf;}");
        html.append(".summary-total td{font-weight:700;background:#fff7bf;}");
        html.append("</style></head><body><div class='page'>");
        html.append("<div class='meta'>Заказчик: ").append(escapeHtml(stringValue(header.get("customer_name")))).append("</div>");
        html.append("<div class='meta'>Объект: ").append(escapeHtml(stringValue(header.get("project_name")))).append(", ").append(escapeHtml(stringValue(header.get("block_name")))).append("</div>");

        for (Map<String, Object> stage : stages) {
            List<Map<String, Object>> subsections = (List<Map<String, Object>>) stage.getOrDefault("subsections", List.of());
            List<Map<String, Object>> materialRows = (List<Map<String, Object>>) stage.getOrDefault("material_rows", List.of());
            List<Map<String, Object>> serviceRows = (List<Map<String, Object>>) stage.getOrDefault("service_rows", List.of());

            html.append("<div class='stage-title'>Расход материалов и услуг на ")
                    .append(escapeHtml(stringValue(stage.get("stage_name"))))
                    .append("</div>");

            appendEstimateStageSectionHtml(html, "Материалы", subsections, materialRows, toDouble(stage.get("material_total_amount")));
            appendEstimateStageSectionHtml(html, "Услуги", subsections, serviceRows, toDouble(stage.get("service_total_amount")));

            html.append("<table><tbody><tr class='total'>")
                    .append("<td style='width:60px'></td>")
                    .append("<td class='left' colspan='").append(Math.max(2, subsections.size() + 2)).append("'>ИТОГО СМР</td>")
                    .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(stage.get("total_amount"))))).append("</td>")
                    .append("</tr></tbody></table>");
        }

        List<Map<String, Object>> summaryRows = (List<Map<String, Object>>) summary.getOrDefault("rows", List.of());
        html.append("<div class='stage-title'>Себестоимость за 1 м2</div>");
        html.append("<div class='meta'>Общая площадь S=").append(escapeHtml(formatSmartNumber(toDouble(summary.get("total_area"))))).append("</div>");
        html.append("<div class='meta'>Продаваемая площадь S=").append(escapeHtml(formatSmartNumber(toDouble(summary.get("sale_area"))))).append("</div>");
        html.append("<table><thead><tr>")
                .append("<th>№</th><th>Наименование</th><th>Ст-ть всего, сом</th><th>Ст-ть всего, $</th><th>С/м2 продаж</th><th>С/м2 общая</th>")
                .append("</tr></thead><tbody>");
        int rowNo = 1;
        for (Map<String, Object> row : summaryRows) {
            html.append("<tr>")
                    .append("<td>").append(rowNo++).append("</td>")
                    .append("<td class='left'>").append(escapeHtml(stringValue(row.get("name")))).append("</td>")
                    .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(row.get("total_amount"))))).append("</td>")
                    .append("<td></td>")
                    .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(row.get("cost_per_sale_area"))))).append("</td>")
                    .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(row.get("cost_per_total_area"))))).append("</td>")
                    .append("</tr>");
        }
        html.append("<tr class='summary-total'>")
                .append("<td></td>")
                .append("<td class='left'>Себестоимость за 1 м2</td>")
                .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(summary.get("total_amount"))))).append("</td>")
                .append("<td></td>")
                .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(summary.get("total_cost_per_sale_area"))))).append("</td>")
                .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(summary.get("total_cost_per_total_area"))))).append("</td>")
                .append("</tr>");
        html.append("</tbody></table>");
        html.append("</div></body></html>");
        return html.toString().getBytes(StandardCharsets.UTF_8);
    }

    private void appendEstimateStageSectionHtml(
            StringBuilder html,
            String title,
            List<Map<String, Object>> subsections,
            List<Map<String, Object>> rows,
            double totalAmount
    ) {
        if (rows == null || rows.isEmpty()) {
            return;
        }

        html.append("<div class='section-title'>").append(escapeHtml(title)).append("</div>");
        html.append("<table><thead><tr>")
                .append("<th>№</th><th>Наименование</th><th>ЕИ</th>");
        for (Map<String, Object> subsection : subsections) {
            html.append("<th>").append(escapeHtml(stringValue(subsection.get("subsection_name")))).append("</th>");
        }
        html.append("<th>Кол-во всего</th><th>Ст-ть ед., сом</th><th>Ст-ть всего, сом</th>")
                .append("</tr></thead><tbody>");

        int rowNo = 1;
        for (Map<String, Object> row : rows) {
            html.append("<tr>")
                    .append("<td>").append(rowNo++).append("</td>")
                    .append("<td class='left'>").append(escapeHtml(stringValue(row.get("item_name")))).append("</td>")
                    .append("<td>").append(escapeHtml(stringValue(row.get("unit_name")))).append("</td>");

            Map subsectionQuantities = (Map) row.getOrDefault("subsection_quantities", Map.of());
            for (Map<String, Object> subsection : subsections) {
                Object value = subsectionQuantities.get(String.valueOf(toInt(subsection.get("subsection_id"))));
                if (value == null) {
                    value = subsectionQuantities.get(toInt(subsection.get("subsection_id")));
                }
                html.append("<td>").append(escapeHtml(formatSmartNumber(toDouble(value)))).append("</td>");
            }

            html.append("<td>").append(escapeHtml(formatSmartNumber(toDouble(row.get("total_quantity"))))).append("</td>")
                    .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(row.get("unit_price"))))).append("</td>")
                    .append("<td>").append(escapeHtml(formatSmartNumber(toDouble(row.get("total_amount"))))).append("</td>")
                    .append("</tr>");
        }

        html.append("<tr class='total'>")
                .append("<td></td>")
                .append("<td class='left' colspan='").append(Math.max(2, subsections.size() + 2)).append("'>Итого ").append(escapeHtml(title.toLowerCase(Locale.ROOT))).append("</td>")
                .append("<td>").append(escapeHtml(formatSmartNumber(totalAmount))).append("</td>")
                .append("</tr>");
        html.append("</tbody></table>");
    }

    private void writeEstimateStageSheet(
            XSSFWorkbook workbook,
            Map<String, Object> stage,
            Map header,
            CellStyle metaStyle,
            CellStyle titleStyle,
            CellStyle headerStyle,
            CellStyle textStyle,
            CellStyle centerStyle,
            CellStyle decimalStyle,
            CellStyle moneyStyle,
            CellStyle totalStyle
    ) {
        String sheetName = sanitizeSheetName(stringValue(stage.get("stage_name")));
        Sheet sheet = workbook.createSheet(sheetName);
        List<Map<String, Object>> subsections = (List<Map<String, Object>>) stage.getOrDefault("subsections", List.of());
        List<Map<String, Object>> materialRows = (List<Map<String, Object>>) stage.getOrDefault("material_rows", List.of());
        List<Map<String, Object>> serviceRows = (List<Map<String, Object>>) stage.getOrDefault("service_rows", List.of());

        int lastColumn = 4 + subsections.size();
        int rowIndex = 0;
        rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "Заказчик: " + stringValue(header.get("customer_name")), metaStyle);
        rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "Объект: " + stringValue(header.get("project_name")) + ", " + stringValue(header.get("block_name")), metaStyle);
        rowIndex++;
        rowIndex = writeMergedValue(sheet, rowIndex, 0, lastColumn, "Расход материалов и услуг на " + stringValue(stage.get("stage_name")), titleStyle);
        rowIndex++;

        Row headerRow = sheet.createRow(rowIndex++);
        writeCell(headerRow, 0, "№", headerStyle);
        writeCell(headerRow, 1, "Наименование", headerStyle);
        writeCell(headerRow, 2, "ЕИ", headerStyle);
        int column = 3;
        for (Map<String, Object> subsection : subsections) {
            writeCell(headerRow, column++, stringValue(subsection.get("subsection_name")), headerStyle);
        }
        writeCell(headerRow, column++, "Кол-во всего", headerStyle);
        writeCell(headerRow, column++, "Ст-ть ед., сом", headerStyle);
        writeCell(headerRow, column, "Ст-ть всего, сом", headerStyle);

        int rowNo = 1;
        if (!materialRows.isEmpty()) {
            Row materialsTitleRow = sheet.createRow(rowIndex++);
            for (int i = 0; i <= lastColumn; i++) {
                writeCell(materialsTitleRow, i, "", headerStyle);
            }
            writeCell(materialsTitleRow, 1, "Материалы", headerStyle);

            for (Map<String, Object> row : materialRows) {
                writeEstimateMatrixRow(sheet.createRow(rowIndex++), rowNo++, row, subsections, textStyle, centerStyle, decimalStyle, moneyStyle);
            }

            Row totalRow = sheet.createRow(rowIndex++);
            for (int i = 0; i <= lastColumn; i++) {
                writeCell(totalRow, i, "", totalStyle);
            }
            writeCell(totalRow, 1, "Итого материалы", totalStyle);
            writeNumericCell(totalRow, 4 + subsections.size(), toDouble(stage.get("material_total_amount")), totalStyle);
        }

        if (!serviceRows.isEmpty()) {
            Row servicesTitleRow = sheet.createRow(rowIndex++);
            for (int i = 0; i <= lastColumn; i++) {
                writeCell(servicesTitleRow, i, "", headerStyle);
            }
            writeCell(servicesTitleRow, 1, "Услуги", headerStyle);

            for (Map<String, Object> row : serviceRows) {
                writeEstimateMatrixRow(sheet.createRow(rowIndex++), rowNo++, row, subsections, textStyle, centerStyle, decimalStyle, moneyStyle);
            }

            Row totalRow = sheet.createRow(rowIndex++);
            for (int i = 0; i <= lastColumn; i++) {
                writeCell(totalRow, i, "", totalStyle);
            }
            writeCell(totalRow, 1, "Итого услуги", totalStyle);
            writeNumericCell(totalRow, 4 + subsections.size(), toDouble(stage.get("service_total_amount")), totalStyle);
        }

        Row grandTotalRow = sheet.createRow(rowIndex);
        for (int i = 0; i <= lastColumn; i++) {
            writeCell(grandTotalRow, i, "", totalStyle);
        }
        writeCell(grandTotalRow, 1, "ИТОГО СМР", totalStyle);
        writeNumericCell(grandTotalRow, 4 + subsections.size(), toDouble(stage.get("total_amount")), totalStyle);

        sheet.setColumnWidth(0, 1800);
        sheet.setColumnWidth(1, 11000);
        sheet.setColumnWidth(2, 2600);
        for (int i = 0; i < subsections.size(); i++) {
            sheet.setColumnWidth(3 + i, 3200);
        }
        sheet.setColumnWidth(3 + subsections.size(), 3600);
        sheet.setColumnWidth(4 + subsections.size(), 3600);
        sheet.setColumnWidth(5 + subsections.size(), 4200);
    }

    private void writeEstimateMatrixRow(
            Row excelRow,
            int rowNo,
            Map<String, Object> row,
            List<Map<String, Object>> subsections,
            CellStyle textStyle,
            CellStyle centerStyle,
            CellStyle decimalStyle,
            CellStyle moneyStyle
    ) {
        writeNumericCell(excelRow, 0, rowNo, centerStyle);
        writeCell(excelRow, 1, stringValue(row.get("item_name")), textStyle);
        writeCell(excelRow, 2, stringValue(row.get("unit_name")), centerStyle);

        Map subsectionQuantities = (Map) row.getOrDefault("subsection_quantities", Map.of());
        int column = 3;
        for (Map<String, Object> subsection : subsections) {
            Object value = subsectionQuantities.get(String.valueOf(toInt(subsection.get("subsection_id"))));
            if (value == null) {
                value = subsectionQuantities.get(toInt(subsection.get("subsection_id")));
            }
            writeSmartNumericCell(excelRow, column++, toDouble(value), centerStyle, decimalStyle);
        }

        writeSmartNumericCell(excelRow, column++, toDouble(row.get("total_quantity")), centerStyle, decimalStyle);
        writeSmartNumericCell(excelRow, column++, toDouble(row.get("unit_price")), centerStyle, moneyStyle);
        writeSmartNumericCell(excelRow, column, toDouble(row.get("total_amount")), centerStyle, moneyStyle);
    }

    private void writeEstimateSummarySheet(
            XSSFWorkbook workbook,
            Map<String, Object> summary,
            Map header,
            CellStyle metaStyle,
            CellStyle titleStyle,
            CellStyle headerStyle,
            CellStyle textStyle,
            CellStyle centerStyle,
            CellStyle moneyStyle,
            CellStyle totalStyle
    ) {
        Sheet sheet = workbook.createSheet("Себестоимость за 1 м2");
        List<Map<String, Object>> rows = (List<Map<String, Object>>) summary.getOrDefault("rows", List.of());

        int rowIndex = 0;
        rowIndex = writeMergedValue(sheet, rowIndex, 0, 5, "УКРУПНЕННЫЙ РАСЧЕТ", titleStyle);
        rowIndex = writeMergedValue(sheet, rowIndex, 0, 5, "по объекту: " + stringValue(header.get("project_name")) + ", " + stringValue(header.get("block_name")), metaStyle);
        rowIndex++;
        rowIndex = writeMergedValue(sheet, rowIndex, 0, 5, "Общая площадь S=" + formatSmartNumber(toDouble(summary.get("total_area"))), metaStyle);
        rowIndex = writeMergedValue(sheet, rowIndex, 0, 5, "Продаваемая площадь S=" + formatSmartNumber(toDouble(summary.get("sale_area"))), metaStyle);
        rowIndex++;

        Row headerRow = sheet.createRow(rowIndex++);
        writeCell(headerRow, 0, "№", headerStyle);
        writeCell(headerRow, 1, "Наименование", headerStyle);
        writeCell(headerRow, 2, "Ст-ть всего, сом", headerStyle);
        writeCell(headerRow, 3, "Ст-ть всего, $", headerStyle);
        writeCell(headerRow, 4, "С/м2 продаж", headerStyle);
        writeCell(headerRow, 5, "С/м2 общая", headerStyle);

        int rowNo = 1;
        for (Map<String, Object> row : rows) {
            Row excelRow = sheet.createRow(rowIndex++);
            writeNumericCell(excelRow, 0, rowNo++, centerStyle);
            writeCell(excelRow, 1, stringValue(row.get("name")), textStyle);
            writeSmartNumericCell(excelRow, 2, toDouble(row.get("total_amount")), centerStyle, moneyStyle);
            writeCell(excelRow, 3, "", centerStyle);
            writeSmartNumericCell(excelRow, 4, toDouble(row.get("cost_per_sale_area")), centerStyle, moneyStyle);
            writeSmartNumericCell(excelRow, 5, toDouble(row.get("cost_per_total_area")), centerStyle, moneyStyle);
        }

        Row totalRow = sheet.createRow(rowIndex);
        for (int i = 0; i <= 5; i++) {
            writeCell(totalRow, i, "", totalStyle);
        }
        writeCell(totalRow, 1, "Себестоимость за 1 м2", totalStyle);
        writeSmartNumericCell(totalRow, 2, toDouble(summary.get("total_amount")), totalStyle, totalStyle);
        writeCell(totalRow, 3, "", totalStyle);
        writeSmartNumericCell(totalRow, 4, toDouble(summary.get("total_cost_per_sale_area")), totalStyle, totalStyle);
        writeSmartNumericCell(totalRow, 5, toDouble(summary.get("total_cost_per_total_area")), totalStyle, totalStyle);

        sheet.setColumnWidth(0, 1800);
        sheet.setColumnWidth(1, 11000);
        sheet.setColumnWidth(2, 4200);
        sheet.setColumnWidth(3, 3600);
        sheet.setColumnWidth(4, 3600);
        sheet.setColumnWidth(5, 3600);
    }

    private String escapeHtml(String value) {
        if (value == null || value.isEmpty()) {
            return "";
        }
        return value
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;");
    }

    private byte[] buildImagePreviewHtml(String base64, String alt) {
        String html = """
                <!doctype html>
                <html>
                <head>
                  <meta charset="UTF-8">
                  <style>
                    html, body { margin: 0; padding: 0; background: #f3f4f6; font-family: Arial, sans-serif; }
                    .wrap { width: 100%%; min-height: 100vh; overflow: auto; box-sizing: border-box; padding: 16px; }
                    .page { width: max-content; min-width: 100%%; }
                    img { display: block; width: auto; max-width: none; background: white; box-shadow: 0 2px 10px rgba(0,0,0,0.08); }
                  </style>
                </head>
                <body>
                  <div class="wrap">
                    <div class="page">
                      <img alt="%s" src="data:image/png;base64,%s" />
                    </div>
                  </div>
                </body>
                </html>
                """.formatted(alt, base64);

        return html.getBytes(StandardCharsets.UTF_8);
    }

    public byte[] generateSalesSummary(ReportData reportData) throws Exception {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> summaryRows = (List<Map<String, Object>>) data.getOrDefault("summary", List.of());
        List<Map<String, Object>> salesRows = (List<Map<String, Object>>) data.getOrDefault("sales", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.NONE);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, BorderStyle.NONE);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, IndexedColors.GREY_25_PERCENT, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle integerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle moneyStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle totalStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, IndexedColors.LIGHT_YELLOW, BorderStyle.THIN);

            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("0.###"));
            moneyStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            totalStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));

            Sheet salesSheet = workbook.createSheet("Продажи");
            int rowIndex = 0;
            rowIndex = writeMergedValue(salesSheet, rowIndex, 0, 22, "Сводка продаж", titleStyle);
            rowIndex = writeMergedValue(salesSheet, rowIndex, 0, 22, salesReportFilterLabel(header), metaStyle);
            rowIndex++;

            String[] salesColumns = {
                    "Тип записи", "Дата", "Менеджер", "Объект", "Блок", "Этаж", "Квартира", "Код квартиры",
                    "Тип лота", "Комнат", "Площадь, м2", "Клиент", "Телефон", "Тип сделки",
                    "Статус договора", "№ договора", "Дата договора", "Стоимость", "Скидка",
                    "Сумма сделки", "Валюта", "Оплачено", "Остаток"
            };
            Row salesHeader = salesSheet.createRow(rowIndex++);
            for (int i = 0; i < salesColumns.length; i++) {
                writeCell(salesHeader, i, salesColumns[i], headerStyle);
            }

            double totalSales = 0;
            double totalPaid = 0;
            double totalDebt = 0;

            for (Map<String, Object> row : salesRows) {
                Row excelRow = salesSheet.createRow(rowIndex++);
                writeCell(excelRow, 0, stringValue(row.get("record_type")), centerStyle);
                writeCell(excelRow, 1, formatIsoDate(row.get("sale_date")), centerStyle);
                writeCell(excelRow, 2, firstNonBlank(row.get("manager_name"), row.get("manager_username")), textStyle);
                writeCell(excelRow, 3, stringValue(row.get("project_name")), textStyle);
                writeCell(excelRow, 4, stringValue(row.get("block_name")), textStyle);
                writeCell(excelRow, 5, salesFloorLabel(row), centerStyle);
                writeCell(excelRow, 6, "№" + stringValue(row.get("unit_number")), centerStyle);
                writeCell(excelRow, 7, stringValue(row.get("plan_code")), centerStyle);
                writeCell(excelRow, 8, salesLotTypeLabel(row.get("lot_type")), centerStyle);
                writeNumericCell(excelRow, 9, toDouble(row.get("rooms")), integerStyle);
                writeSmartNumericCell(excelRow, 10, toDouble(row.get("area_total")), integerStyle, decimalStyle);
                writeCell(excelRow, 11, stringValue(row.get("client_name")), textStyle);
                writeCell(excelRow, 12, stringValue(row.get("client_phone")), centerStyle);
                writeCell(excelRow, 13, stringValue(row.get("deal_type_name")), centerStyle);
                writeCell(excelRow, 14, stringValue(row.get("deal_status_name")), centerStyle);
                writeCell(excelRow, 15, firstNonBlank(row.get("contract_number"), firstNonBlank(row.get("deal_number"), row.get("deal_id"))), centerStyle);
                writeCell(excelRow, 16, formatIsoDate(row.get("contract_date")), centerStyle);
                writeNumericCell(excelRow, 17, toDouble(row.get("total_amount")), moneyStyle);
                writeNumericCell(excelRow, 18, toDouble(row.get("discount_amount")), moneyStyle);
                writeNumericCell(excelRow, 19, toDouble(firstNonBlank(row.get("final_amount"), firstNonBlank(row.get("total_amount"), row.get("unit_price")))), moneyStyle);
                writeCell(excelRow, 20, firstNonBlank(row.get("currency_code"), "KGS"), centerStyle);
                writeNumericCell(excelRow, 21, toDouble(row.get("paid_amount")), moneyStyle);
                writeNumericCell(excelRow, 22, toDouble(row.get("debt_amount")), moneyStyle);

                totalSales += toDouble(firstNonBlank(row.get("final_amount"), firstNonBlank(row.get("total_amount"), row.get("unit_price"))));
                totalPaid += toDouble(row.get("paid_amount"));
                totalDebt += toDouble(row.get("debt_amount"));
            }

            Row totalRow = salesSheet.createRow(rowIndex);
            writeCell(totalRow, 0, "Итого", totalStyle);
            for (int i = 1; i < 19; i++) {
                writeCell(totalRow, i, "", totalStyle);
            }
            writeNumericCell(totalRow, 19, totalSales, totalStyle);
            writeCell(totalRow, 20, "", totalStyle);
            writeNumericCell(totalRow, 21, totalPaid, totalStyle);
            writeNumericCell(totalRow, 22, totalDebt, totalStyle);

            Sheet summarySheet = workbook.createSheet("Итого менеджеры");
            rowIndex = 0;
            rowIndex = writeMergedValue(summarySheet, rowIndex, 0, 6, "Итого по менеджерам", titleStyle);
            rowIndex = writeMergedValue(summarySheet, rowIndex, 0, 6, salesReportFilterLabel(header), metaStyle);
            rowIndex++;

            String[] summaryColumns = {
                    "Менеджер", "Продаж", "Броней", "Площадь, м2", "Сумма сделок", "Оплачено", "Остаток"
            };
            Row summaryHeader = summarySheet.createRow(rowIndex++);
            for (int i = 0; i < summaryColumns.length; i++) {
                writeCell(summaryHeader, i, summaryColumns[i], headerStyle);
            }

            double totalUnits = 0;
            double totalReservations = 0;
            double totalArea = 0;
            double totalManagerSales = 0;
            double totalManagerPaid = 0;
            double totalManagerDebt = 0;

            for (Map<String, Object> row : summaryRows) {
                Row excelRow = summarySheet.createRow(rowIndex++);
                writeCell(excelRow, 0, stringValue(row.get("manager_name")), textStyle);
                writeNumericCell(excelRow, 1, toDouble(row.get("sold_units")), integerStyle);
                writeNumericCell(excelRow, 2, toDouble(row.get("reserved_units")), integerStyle);
                writeSmartNumericCell(excelRow, 3, toDouble(row.get("total_area")), integerStyle, decimalStyle);
                writeNumericCell(excelRow, 4, toDouble(row.get("total_amount")), moneyStyle);
                writeNumericCell(excelRow, 5, toDouble(row.get("paid_amount")), moneyStyle);
                writeNumericCell(excelRow, 6, toDouble(row.get("debt_amount")), moneyStyle);

                totalUnits += toDouble(row.get("sold_units"));
                totalReservations += toDouble(row.get("reserved_units"));
                totalArea += toDouble(row.get("total_area"));
                totalManagerSales += toDouble(row.get("total_amount"));
                totalManagerPaid += toDouble(row.get("paid_amount"));
                totalManagerDebt += toDouble(row.get("debt_amount"));
            }

            Row summaryTotalRow = summarySheet.createRow(rowIndex);
            writeCell(summaryTotalRow, 0, "Итого", totalStyle);
            writeNumericCell(summaryTotalRow, 1, totalUnits, totalStyle);
            writeNumericCell(summaryTotalRow, 2, totalReservations, totalStyle);
            writeSmartNumericCell(summaryTotalRow, 3, totalArea, totalStyle, totalStyle);
            writeNumericCell(summaryTotalRow, 4, totalManagerSales, totalStyle);
            writeNumericCell(summaryTotalRow, 5, totalManagerPaid, totalStyle);
            writeNumericCell(summaryTotalRow, 6, totalManagerDebt, totalStyle);

            autosizeColumns(salesSheet, salesColumns.length);
            autosizeColumns(summarySheet, summaryColumns.length);
            workbook.write(out);
            return out.toByteArray();
        }
    }

    public byte[] generatePaymentsCashBook(ReportData reportData) throws Exception {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.NONE);
            CellStyle subtitleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.NONE);
            CellStyle metaStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.NONE);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle moneyStyle = createStyle(workbook, false, HorizontalAlignment.RIGHT, null, BorderStyle.THIN);
            CellStyle totalLabelStyle = createStyle(workbook, true, HorizontalAlignment.RIGHT, IndexedColors.GREY_25_PERCENT, BorderStyle.THIN);
            CellStyle totalMoneyStyle = createStyle(workbook, true, HorizontalAlignment.RIGHT, IndexedColors.GREY_25_PERCENT, BorderStyle.THIN);

            moneyStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
            totalMoneyStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));

            Sheet sheet = workbook.createSheet("Кассовая книга");
            sheet.setDisplayGridlines(true);

            sheet.setColumnWidth(0, 10000);
            sheet.setColumnWidth(1, 10000);
            sheet.setColumnWidth(2, 10500);
            sheet.setColumnWidth(3, 4200);
            sheet.setColumnWidth(4, 4200);

            int rowIndex = 1;
            String dateFrom = formatIsoDate(header.get("date_from"));
            String dateTo = formatIsoDate(header.get("date_to"));
            String periodLabel = firstNonBlank(dateFrom, "...") + " по " + firstNonBlank(dateTo, "...");
            double openingBalance = toDouble(header.get("opening_balance"));
            double closingBalance = toDouble(header.get("closing_balance"));

            rowIndex = writeMergedValue(sheet, rowIndex, 0, 4, "КАССА за период с " + periodLabel, titleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 4, "Вкладной лист кассовой книги", subtitleStyle);
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 4, paymentsCashBookFilterLabel(header), metaStyle);

            Row openingSummaryRow = sheet.createRow(rowIndex++);
            writeCell(openingSummaryRow, 0, "Остаток на начало", totalLabelStyle);
            writeCell(openingSummaryRow, 1, "", totalLabelStyle);
            writeCell(openingSummaryRow, 2, "", totalLabelStyle);
            writeNumericCell(openingSummaryRow, 3, openingBalance, totalMoneyStyle);
            writeCell(openingSummaryRow, 4, "KGS", totalLabelStyle);
            mergeCellsIfNeeded(sheet, openingSummaryRow.getRowNum(), 0, 2);

            Row closingSummaryRow = sheet.createRow(rowIndex++);
            writeCell(closingSummaryRow, 0, "Остаток на конец", totalLabelStyle);
            writeCell(closingSummaryRow, 1, "", totalLabelStyle);
            writeCell(closingSummaryRow, 2, "", totalLabelStyle);
            writeNumericCell(closingSummaryRow, 3, closingBalance, totalMoneyStyle);
            writeCell(closingSummaryRow, 4, "KGS", totalLabelStyle);
            mergeCellsIfNeeded(sheet, closingSummaryRow.getRowNum(), 0, 2);

            rowIndex++;

            String[] columns = {
                    "Номер документа",
                    "От кого получено или кому выдано",
                    "Данные платежа",
                    "Приход (KGS)",
                    "Расход (KGS)"
            };
            Row headerRow = sheet.createRow(rowIndex++);
            headerRow.setHeightInPoints(42);
            for (int i = 0; i < columns.length; i++) {
                writeCell(headerRow, i, columns[i], headerStyle);
            }

            Row numberRow = sheet.createRow(rowIndex++);
            for (int i = 0; i < columns.length; i++) {
                writeNumericCell(numberRow, i, i + 1, centerStyle);
            }

            for (Map<String, Object> row : rows) {
                Row excelRow = sheet.createRow(rowIndex++);
                String paymentTypeCode = stringValue(row.get("payment_type_code"));
                boolean isIncome = "income".equals(paymentTypeCode);
                String orderName = isIncome ? "Приходный кассовый ордер" : "Расходный кассовый ордер";
                String documentNumber = firstNonBlank(row.get("document_number"), firstNonBlank(row.get("external_number"), row.get("id")));
                String operationDate = formatIsoDate(row.get("operation_date"));
                String documentLabel = orderName + " № " + documentNumber + (operationDate.isBlank() ? "" : " от " + operationDate);
                String counterparty = firstNonBlank(row.get("counterparty_name"), firstNonBlank(row.get("article_name"), row.get("title")));
                String paymentDetails = paymentsCashBookDetails(row);

                writeCell(excelRow, 0, documentLabel, textStyle);
                writeCell(excelRow, 1, counterparty, textStyle);
                writeCell(excelRow, 2, paymentDetails, textStyle);
                writeNumericCell(excelRow, 3, toDouble(row.get("income_kgs")), moneyStyle);
                writeNumericCell(excelRow, 4, toDouble(row.get("expense_kgs")), moneyStyle);
            }

            Row totalRow = sheet.createRow(rowIndex++);
            writeCell(totalRow, 0, "", totalLabelStyle);
            writeCell(totalRow, 1, "Итого за период", totalLabelStyle);
            writeCell(totalRow, 2, "", totalLabelStyle);
            writeNumericCell(totalRow, 3, toDouble(header.get("total_income")), totalMoneyStyle);
            writeNumericCell(totalRow, 4, toDouble(header.get("total_expense")), totalMoneyStyle);

            workbook.write(out);
            return out.toByteArray();
        }
    }

    public byte[] generatePaymentsCashBook(ReportData reportData, String format) throws Exception {
        if ("html".equalsIgnoreCase(format)) {
            return generatePaymentsCashBookHtml(reportData);
        }

        return generatePaymentsCashBook(reportData);
    }

    private byte[] generatePaymentsCashBookHtml(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        String dateFrom = formatIsoDate(header.get("date_from"));
        String dateTo = formatIsoDate(header.get("date_to"));
        String periodLabel = firstNonBlank(dateFrom, "...") + " по " + firstNonBlank(dateTo, "...");
        double openingBalance = toDouble(header.get("opening_balance"));
        double closingBalance = toDouble(header.get("closing_balance"));

        StringBuilder html = new StringBuilder();
        html.append("""
                <!doctype html>
                <html>
                <head>
                  <meta charset="UTF-8">
                  <style>
                    html, body { margin: 0; padding: 0; background: #f8fafc; color: #0f172a; font-family: Arial, sans-serif; }
                    .wrap { padding: 18px; min-width: 980px; box-sizing: border-box; }
                    .page { background: #fff; border: 1px solid #cbd5e1; box-shadow: 0 8px 24px rgba(15,23,42,.08); }
                    .title { text-align: center; font-size: 18px; font-weight: 700; padding: 12px 10px 2px; }
                    .subtitle { text-align: center; font-size: 13px; font-weight: 700; padding-bottom: 8px; }
                    .meta { border-top: 1px solid #0f172a; border-bottom: 1px solid #0f172a; padding: 6px 10px; font-size: 13px; }
                    .balances { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; padding: 10px; }
                    .balance { border: 1px solid #94a3b8; background: #f1f5f9; padding: 8px 10px; display: flex; justify-content: space-between; gap: 12px; font-weight: 700; }
                    table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 12px; }
                    th, td { border: 1px solid #0f172a; padding: 6px 7px; vertical-align: middle; }
                    th { text-align: center; font-weight: 700; background: #f8fafc; }
                    .num { text-align: right; white-space: nowrap; }
                    .center { text-align: center; }
                    .total td { background: #d9d9d9; font-weight: 700; }
                    .muted { color: #64748b; }
                    .doc { width: 30%; }
                    .party { width: 27%; }
                    .details { width: 22%; }
                    .money { width: 10.5%; }
                  </style>
                </head>
                <body>
                <div class="wrap">
                  <div class="page">
                """);

        html.append("<div class=\"title\">КАССА за период с ")
                .append(escapeHtml(periodLabel))
                .append("</div>");
        html.append("<div class=\"subtitle\">Вкладной лист кассовой книги</div>");
        html.append("<div class=\"meta\">")
                .append(escapeHtml(paymentsCashBookFilterLabel(header)))
                .append("</div>");
        html.append("<div class=\"balances\">")
                .append("<div class=\"balance\"><span>Остаток на начало</span><span>")
                .append(escapeHtml(formatMoneyForHtml(openingBalance)))
                .append(" KGS</span></div>")
                .append("<div class=\"balance\"><span>Остаток на конец</span><span>")
                .append(escapeHtml(formatMoneyForHtml(closingBalance)))
                .append(" KGS</span></div>")
                .append("</div>");

        html.append("""
                    <table>
                      <thead>
                        <tr>
                          <th class="doc">Номер документа</th>
                          <th class="party">От кого получено или кому выдано</th>
                          <th class="details">Данные платежа</th>
                          <th class="money">Приход (KGS)</th>
                          <th class="money">Расход (KGS)</th>
                        </tr>
                        <tr class="muted">
                          <th>1</th>
                          <th>2</th>
                          <th>3</th>
                          <th>4</th>
                          <th>5</th>
                        </tr>
                      </thead>
                      <tbody>
                """);

        for (Map<String, Object> row : rows) {
            String paymentTypeCode = stringValue(row.get("payment_type_code"));
            boolean isIncome = "income".equals(paymentTypeCode);
            String orderName = isIncome ? "Приходный кассовый ордер" : "Расходный кассовый ордер";
            String documentNumber = firstNonBlank(row.get("document_number"), firstNonBlank(row.get("external_number"), row.get("id")));
            String operationDate = formatIsoDate(row.get("operation_date"));
            String documentLabel = orderName + " № " + documentNumber + (operationDate.isBlank() ? "" : " от " + operationDate);
            String counterparty = firstNonBlank(row.get("counterparty_name"), firstNonBlank(row.get("article_name"), row.get("title")));
            String paymentDetails = paymentsCashBookDetails(row);

            html.append("<tr>")
                    .append("<td>").append(escapeHtml(documentLabel)).append("</td>")
                    .append("<td>").append(escapeHtml(counterparty)).append("</td>")
                    .append("<td>").append(escapeHtml(paymentDetails)).append("</td>")
                    .append("<td class=\"num\">").append(escapeHtml(formatMoneyForHtml(toDouble(row.get("income_kgs"))))).append("</td>")
                    .append("<td class=\"num\">").append(escapeHtml(formatMoneyForHtml(toDouble(row.get("expense_kgs"))))).append("</td>")
                    .append("</tr>");
        }

        html.append("<tr class=\"total\">")
                .append("<td colspan=\"3\">Итого за период</td>")
                .append("<td class=\"num\">").append(escapeHtml(formatMoneyForHtml(toDouble(header.get("total_income"))))).append("</td>")
                .append("<td class=\"num\">").append(escapeHtml(formatMoneyForHtml(toDouble(header.get("total_expense"))))).append("</td>")
                .append("</tr>");
        html.append("""
                      </tbody>
                    </table>
                  </div>
                </div>
                </body>
                </html>
                """);

        return html.toString().getBytes(StandardCharsets.UTF_8);
    }

    private String formatMoneyForHtml(double value) {
        return String.format(Locale.US, "%,.2f", value).replace(",", " ").replace(".", ",");
    }

    public byte[] generateSalesPaymentSchedule(ReportData reportData) throws Exception {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.NONE);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, BorderStyle.NONE);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, IndexedColors.GREY_25_PERCENT, BorderStyle.THIN);
            CellStyle textStyle = createStyle(workbook, false, HorizontalAlignment.LEFT, null, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle integerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle moneyStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle overdueStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, IndexedColors.ROSE, BorderStyle.THIN);
            CellStyle paidStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, IndexedColors.LIGHT_GREEN, BorderStyle.THIN);
            CellStyle partialStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, IndexedColors.LIGHT_YELLOW, BorderStyle.THIN);
            CellStyle totalStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, IndexedColors.LIGHT_YELLOW, BorderStyle.THIN);

            moneyStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            overdueStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            paidStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            partialStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            totalStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));

            Sheet sheet = workbook.createSheet("График платежей");
            int rowIndex = 0;
            String contractTitle = firstNonBlank(header.get("contract_number"), header.get("deal_number"));
            rowIndex = writeMergedValue(
                    sheet,
                    rowIndex,
                    0,
                    15,
                    contractTitle.isBlank() ? "График платежей отдела продаж" : "График платежей по договору №" + contractTitle,
                    titleStyle
            );
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 15, salesReportFilterLabel(header), metaStyle);
            rowIndex++;

            String[] columns = {
                    "Объект", "Блок", "Этаж", "Лот", "Код", "Клиент", "Телефон", "Договор",
                    "№ платежа", "Плановая дата", "План", "Оплачено", "Остаток", "Статус",
                    "Просрочка, дней", "Связанные платежи"
            };
            Row headerRow = sheet.createRow(rowIndex++);
            for (int i = 0; i < columns.length; i++) {
                writeCell(headerRow, i, columns[i], headerStyle);
            }

            for (Map<String, Object> row : rows) {
                String statusCode = stringValue(row.get("status_code"));
                CellStyle statusStyle = switch (statusCode) {
                    case "paid_on_time", "paid_late" -> paidStyle;
                    case "partial" -> partialStyle;
                    case "overdue" -> overdueStyle;
                    default -> centerStyle;
                };

                Row excelRow = sheet.createRow(rowIndex++);
                writeCell(excelRow, 0, stringValue(row.get("project_name")), textStyle);
                writeCell(excelRow, 1, stringValue(row.get("block_name")), textStyle);
                writeCell(excelRow, 2, salesFloorLabel(row), centerStyle);
                writeCell(excelRow, 3, "№" + stringValue(row.get("unit_number")), centerStyle);
                writeCell(excelRow, 4, stringValue(row.get("plan_code")), centerStyle);
                writeCell(excelRow, 5, stringValue(row.get("client_name")), textStyle);
                writeCell(excelRow, 6, stringValue(row.get("client_phone")), centerStyle);
                writeCell(excelRow, 7, firstNonBlank(row.get("contract_number"), firstNonBlank(row.get("deal_number"), row.get("deal_id"))), centerStyle);
                writeNumericCell(excelRow, 8, toDouble(row.get("payment_no")), integerStyle);
                writeCell(excelRow, 9, formatIsoDate(row.get("planned_date")), centerStyle);
                writeNumericCell(excelRow, 10, toDouble(row.get("planned_amount")), moneyStyle);
                writeNumericCell(excelRow, 11, toDouble(row.get("paid_amount")), moneyStyle);
                writeNumericCell(excelRow, 12, toDouble(row.get("remaining_amount")), moneyStyle);
                writeCell(excelRow, 13, stringValue(row.get("status_name")), statusStyle);
                writeNumericCell(excelRow, 14, toDouble(row.get("overdue_days")), statusStyle);
                writeCell(excelRow, 15, stringValue(row.get("linked_payments")), textStyle);
            }

            Row totalRow = sheet.createRow(rowIndex);
            writeCell(totalRow, 0, "Итого", totalStyle);
            for (int i = 1; i < 10; i++) {
                writeCell(totalRow, i, "", totalStyle);
            }
            writeNumericCell(totalRow, 10, toDouble(header.get("total_planned")), totalStyle);
            writeNumericCell(totalRow, 11, toDouble(header.get("total_paid")), totalStyle);
            writeNumericCell(totalRow, 12, toDouble(header.get("total_remaining")), totalStyle);
            for (int i = 13; i < columns.length; i++) {
                writeCell(totalRow, i, "", totalStyle);
            }

            autosizeColumns(sheet, columns.length);
            workbook.write(out);
            return out.toByteArray();
        }
    }

    public byte[] generateSalesMonthlyPlan(ReportData reportData, String format) throws Exception {
        if ("html".equalsIgnoreCase(format)) {
            return generateSalesMonthlyPlanHtml(reportData);
        }

        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());

        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            CellStyle titleStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, null, BorderStyle.NONE);
            CellStyle metaStyle = createStyle(workbook, true, HorizontalAlignment.LEFT, null, BorderStyle.NONE);
            CellStyle headerStyle = createStyle(workbook, true, HorizontalAlignment.CENTER, IndexedColors.GREY_25_PERCENT, BorderStyle.THIN);
            CellStyle centerStyle = createStyle(workbook, false, HorizontalAlignment.CENTER, null, BorderStyle.THIN);
            CellStyle decimalStyle = createStyle(workbook, false, HorizontalAlignment.RIGHT, null, BorderStyle.THIN);
            CellStyle moneyStyle = createStyle(workbook, false, HorizontalAlignment.RIGHT, null, BorderStyle.THIN);
            CellStyle kgsMoneyStyle = createStyle(workbook, false, HorizontalAlignment.RIGHT, null, BorderStyle.THIN);
            CellStyle totalStyle = createStyle(workbook, true, HorizontalAlignment.RIGHT, IndexedColors.LIGHT_YELLOW, BorderStyle.THIN);
            CellStyle totalMoneyStyle = createStyle(workbook, true, HorizontalAlignment.RIGHT, IndexedColors.LIGHT_YELLOW, BorderStyle.THIN);
            CellStyle totalKgsMoneyStyle = createStyle(workbook, true, HorizontalAlignment.RIGHT, IndexedColors.LIGHT_YELLOW, BorderStyle.THIN);

            decimalStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.###"));
            moneyStyle.setDataFormat(workbook.createDataFormat().getFormat("$#,##0.00"));
            kgsMoneyStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
            totalMoneyStyle.setDataFormat(workbook.createDataFormat().getFormat("$#,##0.00"));
            totalKgsMoneyStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));

            Sheet sheet = workbook.createSheet("График продаж");
            int rowIndex = 0;
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 8, "ПОМЕСЯЧНЫЙ ГРАФИК ПРОДАЖ", titleStyle);
            String scope = "Объект: " + stringValue(header.get("project_name"));
            if (!stringValue(header.get("block_name")).isBlank()) {
                scope += " | Блок: " + stringValue(header.get("block_name"));
            }
            scope += " | Период: " + stringValue(header.get("date_from")) + " - " + stringValue(header.get("date_to")) + " | Валюты: USD / KGS";
            rowIndex = writeMergedValue(sheet, rowIndex, 0, 8, scope, metaStyle);
            rowIndex++;

            Row groupHeaderRow = sheet.createRow(rowIndex++);
            Row valueHeaderRow = sheet.createRow(rowIndex++);
            writeCell(groupHeaderRow, 0, "Месяц", headerStyle);
            writeCell(valueHeaderRow, 0, "", headerStyle);
            mergeCellsIfNeeded(sheet, groupHeaderRow.getRowNum(), 0, valueHeaderRow.getRowNum(), 0);

            String[] groups = {"Лоты", "м²", "Поступления USD", "Поступления KGS"};
            for (int groupIndex = 0; groupIndex < groups.length; groupIndex++) {
                int firstColumn = 1 + groupIndex * 2;
                writeCell(groupHeaderRow, firstColumn, groups[groupIndex], headerStyle);
                writeCell(groupHeaderRow, firstColumn + 1, "", headerStyle);
                mergeCellsIfNeeded(sheet, groupHeaderRow.getRowNum(), firstColumn, firstColumn + 1);
                writeCell(valueHeaderRow, firstColumn, "Потребность", headerStyle);
                writeCell(valueHeaderRow, firstColumn + 1, "Факт", headerStyle);
            }

            for (Map<String, Object> row : rows) {
                Row excelRow = sheet.createRow(rowIndex++);
                writeCell(excelRow, 0, stringValue(row.get("month_name")), centerStyle);
                writeNumericCell(excelRow, 1, toDouble(row.get("planned_units")), centerStyle);
                writeNumericCell(excelRow, 2, toDouble(row.get("actual_units")), centerStyle);
                writeNumericCell(excelRow, 3, toDouble(row.get("planned_area")), decimalStyle);
                writeNumericCell(excelRow, 4, toDouble(row.get("actual_area")), decimalStyle);
                writeNumericCell(excelRow, 5, toDouble(row.get("planned_cash_amount_usd")), moneyStyle);
                writeNumericCell(excelRow, 6, toDouble(row.get("actual_cash_amount_usd")), moneyStyle);
                writeNumericCell(excelRow, 7, toDouble(row.get("planned_cash_amount_kgs")), kgsMoneyStyle);
                writeNumericCell(excelRow, 8, toDouble(row.get("actual_cash_amount_kgs")), kgsMoneyStyle);
            }

            Row totalRow = sheet.createRow(rowIndex);
            writeCell(totalRow, 0, "Итого", totalStyle);
            writeNumericCell(totalRow, 1, toDouble(header.get("total_planned_units")), totalStyle);
            writeNumericCell(totalRow, 2, toDouble(header.get("total_actual_units")), totalStyle);
            writeNumericCell(totalRow, 3, toDouble(header.get("total_planned_area")), totalStyle);
            writeNumericCell(totalRow, 4, toDouble(header.get("total_actual_area")), totalStyle);
            writeNumericCell(totalRow, 5, toDouble(header.get("total_planned_cash_amount_usd")), totalMoneyStyle);
            writeNumericCell(totalRow, 6, toDouble(header.get("total_actual_cash_amount_usd")), totalMoneyStyle);
            writeNumericCell(totalRow, 7, toDouble(header.get("total_planned_cash_amount_kgs")), totalKgsMoneyStyle);
            writeNumericCell(totalRow, 8, toDouble(header.get("total_actual_cash_amount_kgs")), totalKgsMoneyStyle);

            autosizeColumns(sheet, 9);
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private byte[] generateSalesMonthlyPlanHtml(ReportData reportData) {
        Map data = reportData.data();
        Map header = reportData.header();
        List<Map<String, Object>> rows = (List<Map<String, Object>>) data.getOrDefault("rows", List.of());
        StringBuilder html = new StringBuilder();
        html.append("<!doctype html><html><head><meta charset='UTF-8'><style>")
                .append("html,body{margin:0;padding:0;background:#f3f4f6;color:#172033}")
                .append("body{font-family:Arial,sans-serif;padding:16px;box-sizing:border-box}")
                .append(".sheet{background:#fff;padding:16px;min-width:980px;box-shadow:0 2px 10px rgba(0,0,0,.08)}")
                .append("h2{margin:0 0 6px;text-align:center;font-size:20px;text-transform:uppercase}")
                .append(".meta{font-weight:700;margin-bottom:16px;white-space:nowrap}")
                .append("table{border-collapse:collapse;width:100%;font-size:12px;table-layout:auto}")
                .append("th,td{border:1px solid #475569;padding:7px 8px;text-align:right;vertical-align:middle}")
                .append("th{background:#d1d5db;text-align:center;font-weight:700;white-space:normal}")
                .append("td:first-child{text-align:left;white-space:nowrap}")
                .append(".money{white-space:nowrap}.total td{background:#fff59d;font-weight:700}")
                .append("</style></head><body><div class='sheet'>")
                .append("<h2>Помесячный график продаж</h2><div class='meta'>Объект: ")
                .append(escapeHtml(stringValue(header.get("project_name"))));
        if (!stringValue(header.get("block_name")).isBlank()) {
            html.append(" | Блок: ").append(escapeHtml(stringValue(header.get("block_name"))));
        }
        html.append(" | Период: ").append(escapeHtml(stringValue(header.get("date_from"))))
                .append(" - ").append(escapeHtml(stringValue(header.get("date_to"))))
                .append(" | Валюты: USD / KGS</div>")
                .append("<table><thead><tr>")
                .append("<th rowspan='2'>Месяц</th>")
                .append("<th colspan='2'>Лоты</th>")
                .append("<th colspan='2'>м²</th>")
                .append("<th colspan='2'>Поступления USD</th>")
                .append("<th colspan='2'>Поступления KGS</th>")
                .append("</tr><tr>")
                .append("<th>Потребность</th><th>Факт</th>")
                .append("<th>Потребность</th><th>Факт</th>")
                .append("<th>Потребность</th><th>Факт</th>")
                .append("<th>Потребность</th><th>Факт</th>")
                .append("</tr></thead><tbody>");

        for (Map<String, Object> row : rows) {
            html.append("<tr><td class='left'>").append(escapeHtml(stringValue(row.get("month_name")))).append("</td>")
                    .append("<td>").append(formatSmartNumber(toDouble(row.get("planned_units")))).append("</td>")
                    .append("<td>").append(formatSmartNumber(toDouble(row.get("actual_units")))).append("</td>")
                    .append("<td>").append(formatSmartNumber(toDouble(row.get("planned_area")))).append("</td>")
                    .append("<td>").append(formatSmartNumber(toDouble(row.get("actual_area")))).append("</td>")
                    .append("<td class='money'>$").append(formatMoneyForHtml(toDouble(row.get("planned_cash_amount_usd")))).append("</td>")
                    .append("<td class='money'>$").append(formatMoneyForHtml(toDouble(row.get("actual_cash_amount_usd")))).append("</td>")
                    .append("<td class='money'>").append(formatMoneyForHtml(toDouble(row.get("planned_cash_amount_kgs")))).append("</td>")
                    .append("<td class='money'>").append(formatMoneyForHtml(toDouble(row.get("actual_cash_amount_kgs")))).append("</td></tr>");
        }

        html.append("<tr class='total'><td>Итого</td>")
                .append("<td>").append(formatSmartNumber(toDouble(header.get("total_planned_units")))).append("</td>")
                .append("<td>").append(formatSmartNumber(toDouble(header.get("total_actual_units")))).append("</td>")
                .append("<td>").append(formatSmartNumber(toDouble(header.get("total_planned_area")))).append("</td>")
                .append("<td>").append(formatSmartNumber(toDouble(header.get("total_actual_area")))).append("</td>")
                .append("<td class='money'>$").append(formatMoneyForHtml(toDouble(header.get("total_planned_cash_amount_usd")))).append("</td>")
                .append("<td class='money'>$").append(formatMoneyForHtml(toDouble(header.get("total_actual_cash_amount_usd")))).append("</td>")
                .append("<td class='money'>").append(formatMoneyForHtml(toDouble(header.get("total_planned_cash_amount_kgs")))).append("</td>")
                .append("<td class='money'>").append(formatMoneyForHtml(toDouble(header.get("total_actual_cash_amount_kgs")))).append("</td>")
                .append("</tr></tbody></table></div></body></html>");
        return html.toString().getBytes(StandardCharsets.UTF_8);
    }

    private void writeSalesUnitsSheet(
            Sheet sheet,
            CellStyle headerStyle,
            CellStyle textStyle,
            CellStyle centerStyle,
            CellStyle integerStyle,
            CellStyle decimalStyle,
            CellStyle moneyStyle,
            List<Map<String, Object>> rows
    ) {
        String[] columns = {
                "Объект", "Блок", "Этаж", "Лот", "Код", "Тип", "Комнат", "Площадь",
                "Статус", "Стоимость", "Валюта", "Поступило", "Долг", "Создан"
        };
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columns.length; i++) {
            writeCell(headerRow, i, columns[i], headerStyle);
        }

        int rowIndex = 1;
        for (Map<String, Object> row : rows) {
            Row excelRow = sheet.createRow(rowIndex++);
            writeCell(excelRow, 0, stringValue(row.get("project_name")), textStyle);
            writeCell(excelRow, 1, stringValue(row.get("block_name")), textStyle);
            writeCell(excelRow, 2, salesFloorLabel(row), centerStyle);
            writeCell(excelRow, 3, "№" + stringValue(row.get("unit_number")), centerStyle);
            writeCell(excelRow, 4, stringValue(row.get("plan_code")), centerStyle);
            writeCell(excelRow, 5, salesLotTypeLabel(row.get("lot_type")), centerStyle);
            writeNumericCell(excelRow, 6, toDouble(row.get("rooms")), integerStyle);
            writeSmartNumericCell(excelRow, 7, toDouble(row.get("area_total")), integerStyle, decimalStyle);
            writeCell(excelRow, 8, stringValue(row.get("status_name")), centerStyle);
            writeNumericCell(excelRow, 9, toDouble(row.get("price_total")), moneyStyle);
            writeCell(excelRow, 10, firstNonBlank(row.get("currency_code"), "KGS"), centerStyle);
            writeNumericCell(excelRow, 11, toDouble(row.get("paid_amount")), moneyStyle);
            writeNumericCell(excelRow, 12, toDouble(row.get("debt_amount")), moneyStyle);
            writeCell(excelRow, 13, formatIsoDate(row.get("created_at")), centerStyle);
        }
    }

    private String salesReportFilterLabel(Map header) {
        String project = stringValue(header.get("project_name"));
        String block = stringValue(header.get("block_name"));
        String contract = firstNonBlank(header.get("contract_number"), header.get("deal_number"));
        String unitNumber = stringValue(header.get("unit_number"));
        String client = stringValue(header.get("client_name"));
        String dateFrom = formatIsoDate(header.get("date_from"));
        String dateTo = formatIsoDate(header.get("date_to"));
        List<String> parts = new ArrayList<>();
        if (!project.isBlank()) parts.add("Объект: " + project);
        if (!block.isBlank()) parts.add("Блок: " + block);
        if (!unitNumber.isBlank()) parts.add("Лот: №" + unitNumber);
        if (!contract.isBlank()) parts.add("Договор: №" + contract);
        if (!client.isBlank()) parts.add("Клиент: " + client);
        if (!dateFrom.isBlank() || !dateTo.isBlank()) parts.add("Период: " + firstNonBlank(dateFrom, "...") + " - " + firstNonBlank(dateTo, "..."));
        parts.add("Сформировано: " + formatIsoDate(header.get("generated_at")));
        return String.join(" | ", parts);
    }

    private String paymentsCashBookFilterLabel(Map header) {
        String project = stringValue(header.get("project_name"));
        String block = stringValue(header.get("block_name"));
        String dateFrom = formatIsoDate(header.get("date_from"));
        String dateTo = formatIsoDate(header.get("date_to"));
        List<String> parts = new ArrayList<>();
        if (!project.isBlank()) parts.add("Объект: " + project);
        if (!block.isBlank()) parts.add("Блок: " + block);
        if (!dateFrom.isBlank() || !dateTo.isBlank()) parts.add("Период: " + firstNonBlank(dateFrom, "...") + " - " + firstNonBlank(dateTo, "..."));
        parts.add("Сформировано: " + formatIsoDate(header.get("generated_at")));
        return String.join(" | ", parts);
    }

    private String paymentsCashBookDetails(Map<String, Object> row) {
        String preparedDetails = stringValue(row.get("payment_details"));
        if (!preparedDetails.isBlank()) {
            return preparedDetails;
        }

        List<String> parts = new ArrayList<>();
        addUniqueNonBlank(parts, row.get("article_name"));
        addUniqueNonBlank(parts, row.get("title"));

        String entityTypeName = stringValue(row.get("entity_type_name"));
        String entityId = stringValue(row.get("entity_id"));
        if (!entityTypeName.isBlank()) {
            addUniqueNonBlank(parts, entityId.isBlank() ? entityTypeName : entityTypeName + " №" + entityId);
        }

        addUniqueNonBlank(parts, row.get("description"));
        addUniqueNonBlank(parts, row.get("comment"));

        return String.join(" — ", parts);
    }

    private void addUniqueNonBlank(List<String> parts, Object value) {
        String text = stringValue(value).trim();
        if (!text.isBlank() && !parts.contains(text)) {
            parts.add(text);
        }
    }

    private String salesFloorLabel(Map<String, Object> row) {
        String name = stringValue(row.get("floor_name"));
        if (!name.isBlank()) {
            return name;
        }
        String number = stringValue(row.get("floor_number"));
        return number.isBlank() ? "" : number + " этаж";
    }

    private String salesLotTypeLabel(Object value) {
        return switch (stringValue(value)) {
            case "apartment" -> "Квартира";
            case "commercial" -> "Помещение";
            case "parking" -> "Паркинг";
            case "storage" -> "Кладовая";
            default -> stringValue(value);
        };
    }

    private void autosizeColumns(Sheet sheet, int columns) {
        for (int i = 0; i < columns; i++) {
            sheet.autoSizeColumn(i);
            int width = sheet.getColumnWidth(i);
            sheet.setColumnWidth(i, Math.min(Math.max(width + 700, 3200), 12000));
        }
    }

    public String buildWorkPerformedFilename(String id, String format, String authorizationHeader) {
        return buildWorkPerformedFilename(getWorkPerformedData(id, authorizationHeader), id, format);
    }

    public String buildMaterialRequestFilename(ReportData reportData, String id) {
        String requestId = firstNonBlank(reportData.header().get("id"), id);
        return "Заявка на материалы №" + requestId + ".xlsx";
    }

    public String buildWorkPerformedFilename(ReportData reportData, String id, String format) {
        Map header = reportData.header();
        String reportId = firstNonBlank(header.get("id"), id);
        String projectName = stringValue(header.get("project_name"));
        String blockName = stringValue(header.get("block_name"));
        return sanitizeFilename("Акт вып. работ №" + reportId + " " + projectName + " " + blockName + "." + format);
    }

    public String buildForm29Filename(ReportData reportData, String format) {
        Map header = reportData.header();
        String projectName = stringValue(header.get("project_name"));
        String blockName = stringValue(header.get("block_name"));
        String periodLabel = stringValue(header.get("period_label")).replace(" г.", "");
        return sanitizeFilename("\u0424\u043e\u0440\u043c\u0430 29 " + projectName + " " + blockName + " " + periodLabel + "." + format);
    }

    public String buildForm2Filename(ReportData reportData, String format) {
        Map header = reportData.header();
        String projectName = stringValue(header.get("project_name"));
        String blockName = stringValue(header.get("block_name"));
        String periodLabel = stringValue(header.get("period_label")).replace(" г.", "").replace(" Рі.", "");
        return sanitizeFilename("Форма 2 " + projectName + " " + blockName + " " + periodLabel + "." + format);
    }

    public String buildForm19Filename(ReportData reportData, String format) {
        Map header = reportData.header();
        String projectName = stringValue(header.get("project_name"));
        String warehouseName = stringValue(header.get("warehouse_name"));
        String periodLabel = stringValue(header.get("period_label")).replace(" г.", "");
        return sanitizeFilename("Ф-19 " + projectName + " " + warehouseName + " " + periodLabel + "." + format);
    }

    public String buildMbpWriteOffFilename(ReportData reportData, String format) {
        Map header = reportData.header();
        String projectName = stringValue(header.get("project_name"));
        String periodLabel = stringValue(header.get("period_label")).replace(" г.", "");
        return sanitizeFilename("Списание МБП " + projectName + " " + periodLabel + "." + format);
    }

    public String buildProjectsOverviewFilename(ReportData reportData, String format) {
        Map header = reportData.header();
        return sanitizeFilename("Сводка по проектам " + formatIsoDate(header.get("report_date")) + "." + format);
    }

    public String buildPaymentCashOrderFilename(ReportData reportData, String format) {
        Map header = reportData.header();
        String orderShortName = firstNonBlank(header.get("order_short_name"), "ПКО");
        String orderNumber = firstNonBlank(header.get("order_number"), header.get("id"));
        String projectName = stringValue(header.get("project_name"));
        return sanitizeFilename(orderShortName + " №" + orderNumber + " " + projectName + "." + format);
    }

    public String buildPaymentsCashBookFilename(ReportData reportData) {
        Map header = reportData.header();
        String projectName = firstNonBlank(header.get("project_name"), "Объект");
        String blockName = stringValue(header.get("block_name"));
        String dateFrom = formatIsoDate(header.get("date_from"));
        String dateTo = formatIsoDate(header.get("date_to"));
        return sanitizeFilename("Кассовая книга " + projectName + " " + blockName + " " + dateFrom + " " + dateTo + ".xlsx");
    }

    public String buildSalesSummaryFilename(ReportData reportData) {
        Map header = reportData.header();
        String projectName = firstNonBlank(header.get("project_name"), "Все объекты");
        String blockName = stringValue(header.get("block_name"));
        return sanitizeFilename("Сводка продаж " + projectName + " " + blockName + ".xlsx");
    }

    public String buildSalesPaymentScheduleFilename(ReportData reportData) {
        Map header = reportData.header();
        String projectName = firstNonBlank(header.get("project_name"), "Все объекты");
        String blockName = stringValue(header.get("block_name"));
        String contract = firstNonBlank(header.get("contract_number"), header.get("deal_number"));
        String unitNumber = stringValue(header.get("unit_number"));
        if (!contract.isBlank()) {
            return sanitizeFilename("График платежей договор №" + contract + " лот №" + unitNumber + ".xlsx");
        }
        return sanitizeFilename("График платежей продаж " + projectName + " " + blockName + ".xlsx");
    }

    public String buildSalesMonthlyPlanFilename(ReportData reportData, String format) {
        Map header = reportData.header();
        String projectName = firstNonBlank(header.get("project_name"), "Объект");
        String blockName = stringValue(header.get("block_name"));
        String period = stringValue(header.get("date_from")) + "-" + stringValue(header.get("date_to"));
        return sanitizeFilename("График продаж " + projectName + " " + blockName + " " + period + "." + format);
    }

    public String buildEstimateStageFilename(ReportData reportData, String format) {
        Map header = reportData.header();
        String projectName = stringValue(header.get("project_name"));
        String blockName = stringValue(header.get("block_name"));
        return sanitizeFilename("Смета " + projectName + " " + blockName + "." + format);
    }

    public String buildScheduleFilename(ReportData reportData, String format) {
        Map header = reportData.header();
        String projectName = stringValue(header.get("project_name"));
        String blockName = stringValue(header.get("block_name"));
        String dateFrom = stringValue(header.get("date_from"));
        String dateTo = stringValue(header.get("date_to"));
        String prefix = firstNonBlank(header.get("filename_prefix"), "График работ");
        return sanitizeFilename(prefix + " " + projectName + " " + blockName + " " + dateFrom + " " + dateTo + "." + format);
    }

    private String scheduleReportTitle(Map header) {
        return firstNonBlank(header.get("report_title"), "КАЛЕНДАРНЫЙ ГРАФИК ПРОИЗВОДСТВА РАБОТ");
    }

    private String scheduleNameColumn(Map header) {
        return firstNonBlank(header.get("name_column"), "Наименование работ");
    }

    private String scheduleSheetName(Map header) {
        return firstNonBlank(header.get("sheet_name"), "График работ");
    }

    private String sanitizeFilename(String filename) {
        return filename
                .replaceAll("[\\\\/:*?\"<>|]", " ")
                .replaceAll("\\s+", " ")
                .trim();
    }

    private String sanitizeSheetName(String name) {
        String sanitized = stringValue(name)
                .replaceAll("[\\\\/?*\\[\\]:]", " ")
                .trim();
        if (sanitized.isBlank()) {
            sanitized = "Лист";
        }
        return sanitized.length() > 31 ? sanitized.substring(0, 31) : sanitized;
    }

    private CellStyle createCashOrderStyle(
            XSSFWorkbook workbook,
            boolean bold,
            int fontSize,
            HorizontalAlignment alignment,
            BorderStyle borderStyle
    ) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(alignment);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        if (borderStyle != null && borderStyle != BorderStyle.NONE) {
            style.setBorderTop(borderStyle);
            style.setBorderRight(borderStyle);
            style.setBorderBottom(borderStyle);
            style.setBorderLeft(borderStyle);
        }

        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) fontSize);
        font.setBold(bold);
        style.setFont(font);
        return style;
    }

    private void writeMerged(Sheet sheet, int rowIndex, int startCol, int endCol, String value, CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        for (int col = startCol; col <= endCol; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellValue(col == startCol ? stringValue(value) : "");
            cell.setCellStyle(style);
        }

        if (endCol > startCol) {
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, startCol, endCol));
        }
    }

    private void setRowHeight(Sheet sheet, int rowIndex, int heightInPoints) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        row.setHeightInPoints(heightInPoints);
    }

    private void addBottomBorder(Sheet sheet, int rowIndex, int startCol, int endCol) {
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, startCol, endCol);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
    }

    private void addRightDashedBorder(Sheet sheet, int rowIndex, int col) {
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, col, col);
        RegionUtil.setBorderRight(BorderStyle.DASHED, region, sheet);
    }

    private String cashOrderAmountToWords(double amount, String currency) {
        if (!"KGS".equalsIgnoreCase(currency)) {
            return formatSmartNumber(amount) + " " + currency;
        }

        BigDecimal rounded = BigDecimal.valueOf(amount).setScale(2, RoundingMode.HALF_UP);
        long soms = rounded.longValue();
        int tyiyn = rounded.remainder(BigDecimal.ONE).movePointRight(2).abs().intValue();
        return capitalize(numberToWordsRu(soms)) + " сом " + String.format("%02d", tyiyn) + " тыйын";
    }

    private String numberToWordsRu(long value) {
        if (value == 0) {
            return "ноль";
        }

        StringBuilder result = new StringBuilder();
        appendGroupRu(result, (int) (value / 1_000_000_000), "миллиард", "миллиарда", "миллиардов", false);
        appendGroupRu(result, (int) ((value / 1_000_000) % 1000), "миллион", "миллиона", "миллионов", false);
        appendGroupRu(result, (int) ((value / 1000) % 1000), "тысяча", "тысячи", "тысяч", true);
        appendGroupRu(result, (int) (value % 1000), "", "", "", false);
        return result.toString().trim().replaceAll("\\s+", " ");
    }

    private void appendGroupRu(
            StringBuilder result,
            int value,
            String oneForm,
            String twoForm,
            String manyForm,
            boolean feminine
    ) {
        if (value == 0) {
            return;
        }

        String[] hundredsWords = {"", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"};
        String[] tensWords = {"", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"};
        String[] teensWords = {"десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"};
        String[] unitsMasculine = {"", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"};
        String[] unitsFeminine = {"", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"};

        int hundreds = value / 100;
        int tensUnits = value % 100;
        int tens = tensUnits / 10;
        int units = tensUnits % 10;

        appendWord(result, hundredsWords[hundreds]);

        if (tensUnits >= 10 && tensUnits <= 19) {
            appendWord(result, teensWords[tensUnits - 10]);
        } else {
            appendWord(result, tensWords[tens]);
            appendWord(result, feminine ? unitsFeminine[units] : unitsMasculine[units]);
        }

        if (!oneForm.isBlank()) {
            appendWord(result, pluralForm(value, oneForm, twoForm, manyForm));
        }
    }

    private String formatSignatureName(Object value) {
        String name = stringValue(value).replace("\"", "").trim();
        if (name.isBlank()) {
            return "";
        }

        String prefix = "";
        if (name.toUpperCase().startsWith("ИП ")) {
            prefix = "ИП ";
            name = name.substring(3).trim();
        }

        String[] parts = name.split("\\s+");
        if (parts.length < 2) {
            return prefix + name;
        }

        String surname;
        StringBuilder initials = new StringBuilder();

        if (isInitial(parts[0]) && parts.length >= 2) {
            surname = parts[1];
            initials.append(normalizeInitial(parts[0]));
            for (int i = 2; i < parts.length; i++) {
                initials.append(normalizeInitial(parts[i]));
            }
        } else {
            surname = parts[0];
            for (int i = 1; i < parts.length; i++) {
                initials.append(normalizeInitial(parts[i]));
            }
        }

        if (initials.isEmpty()) {
            return prefix + surname;
        }
        return prefix + surname + " " + initials;
    }

    private boolean isInitial(String value) {
        return value != null && value.matches("[\\p{L}]\\.?[\\p{L}]?\\.?");
    }

    private String normalizeInitial(String value) {
        if (value == null || value.isBlank()) {
            return "";
        }
        String cleaned = value.replace(".", "");
        if (cleaned.isBlank()) {
            return "";
        }
        StringBuilder result = new StringBuilder();
        for (int i = 0; i < cleaned.length(); i++) {
            result.append(cleaned.charAt(i)).append('.');
            if (i == 1) {
                break;
            }
        }
        return result.toString();
    }

    public record ReportData(Map data, Map header) {
    }

    private record ScheduleSlot(LocalDate start, LocalDate end, String yearLabel, String monthLabel, String slotLabel) {
    }

    private record ScheduleRowData(int rowNo, String label, List<Boolean> filled) {
    }

    private record ScheduleRenderModel(
            String projectName,
            String blockName,
            String customerName,
            String reportTitle,
            String nameColumn,
            String sheetName,
            List<ScheduleSlot> slots,
            List<ScheduleRowData> rows
    ) {
    }

    private int writeMergedValue(Sheet sheet, int rowIndex, int startCol, int endCol, String value, CellStyle style) {
        Row row = sheet.createRow(rowIndex);
        for (int col = startCol; col <= endCol; col++) {
            writeCell(row, col, "", style);
        }
        writeCell(row, startCol, value, style);
        mergeCellsIfNeeded(sheet, rowIndex, startCol, endCol);
        return rowIndex + 1;
    }

    private void mergeCellsIfNeeded(Sheet sheet, int rowIndex, int startCol, int endCol) {
        if (endCol <= startCol) {
            return;
        }
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, startCol, endCol);
        sheet.addMergedRegion(region);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
    }

    private void mergeCellsIfNeeded(Sheet sheet, int startRow, int startCol, int endRow, int endCol) {
        if (endRow <= startRow && endCol <= startCol) {
            return;
        }
        CellRangeAddress region = new CellRangeAddress(startRow, endRow, startCol, endCol);
        sheet.addMergedRegion(region);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
    }

    private int writeTotalsRow(
            Sheet sheet,
            int rowIndex,
            List<Map<String, Object>> materials,
            Map<Integer, Map<String, Object>> totalsByMaterial,
            String label,
            String field,
            CellStyle totalIntegerStyle,
            CellStyle totalDecimalStyle,
            CellStyle additionalTotalIntegerStyle,
            CellStyle additionalTotalDecimalStyle,
            CellStyle estimateTotalIntegerStyle,
            CellStyle estimateTotalDecimalStyle,
            CellStyle labelStyle,
            CellStyle additionalLabelStyle,
            CellStyle estimateLabelStyle
    ) {
        Row row = sheet.createRow(rowIndex);
        writeCell(row, 0, label, labelStyle);

        int materialColumn = 3;
        writeCell(row, 1, "", labelStyle);
        writeCell(row, 2, "", labelStyle);
        for (Map<String, Object> material : materials) {
            Map<String, Object> total = totalsByMaterial.get(toInt(material.get("material_id")));
            double value = 0;
            if (total != null) {
                if ("deviation_quantity".equals(field)) {
                    value = toDouble(total.get(field));
                } else {
                    value = toDouble(total.get(field));
                }
            }
            CellStyle materialLabelStyle = resolveMaterialCenterStyle(
                    material,
                    labelStyle,
                    additionalLabelStyle,
                    estimateLabelStyle
            );
            CellStyle materialTotalIntegerStyle = resolveMaterialIntegerStyle(
                    material,
                    totalIntegerStyle,
                    additionalTotalIntegerStyle,
                    estimateTotalIntegerStyle
            );
            CellStyle materialTotalDecimalStyle = resolveMaterialDecimalStyle(
                    material,
                    totalDecimalStyle,
                    additionalTotalDecimalStyle,
                    estimateTotalDecimalStyle
            );
            writeCell(row, materialColumn, "Х", materialLabelStyle);
            if ("deviation_quantity".equals(field)) {
                writeCell(row, materialColumn + 1, formatSignedSmartNumber(value), materialLabelStyle);
            } else {
                writeSmartNumericCell(row, materialColumn + 1, value, materialTotalIntegerStyle, materialTotalDecimalStyle);
            }
            materialColumn += 2;
        }

        return rowIndex + 1;
    }

    private CellStyle resolveMaterialHeaderStyle(
            Map<String, Object> material,
            CellStyle baseStyle,
            CellStyle additionalStyle,
            CellStyle estimateStyle
    ) {
        return resolveMaterialStyle(material, baseStyle, additionalStyle, estimateStyle);
    }

    private CellStyle resolveMaterialCenterStyle(
            Map<String, Object> material,
            CellStyle baseStyle,
            CellStyle additionalStyle,
            CellStyle estimateStyle
    ) {
        return resolveMaterialStyle(material, baseStyle, additionalStyle, estimateStyle);
    }

    private CellStyle resolveMaterialIntegerStyle(
            Map<String, Object> material,
            CellStyle baseStyle,
            CellStyle additionalStyle,
            CellStyle estimateStyle
    ) {
        return resolveMaterialStyle(material, baseStyle, additionalStyle, estimateStyle);
    }

    private CellStyle resolveMaterialDecimalStyle(
            Map<String, Object> material,
            CellStyle baseStyle,
            CellStyle additionalStyle,
            CellStyle estimateStyle
    ) {
        return resolveMaterialStyle(material, baseStyle, additionalStyle, estimateStyle);
    }

    private CellStyle resolveMaterialStyle(
            Map<String, Object> material,
            CellStyle baseStyle,
            CellStyle additionalStyle,
            CellStyle estimateStyle
    ) {
        boolean hasFact = Boolean.TRUE.equals(material.get("has_fact"));
        boolean isAdditional = Boolean.TRUE.equals(material.get("is_additional"));
        boolean fromEstimate = Boolean.TRUE.equals(material.get("from_estimate"));

        if (hasFact && isAdditional) {
            return additionalStyle;
        }
        if (!hasFact && fromEstimate) {
            return estimateStyle;
        }
        return baseStyle;
    }

    private CellStyle createStyle(
            XSSFWorkbook workbook,
            boolean bold,
            HorizontalAlignment alignment,
            IndexedColors fillColor,
            BorderStyle borderStyle
    ) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(alignment);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        if (fillColor != null && fillColor != IndexedColors.WHITE) {
            style.setFillForegroundColor(fillColor.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        if (borderStyle != null) {
            style.setBorderTop(borderStyle);
            style.setBorderRight(borderStyle);
            style.setBorderBottom(borderStyle);
            style.setBorderLeft(borderStyle);
        }

        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setBold(bold);
        style.setFont(font);
        return style;
    }

    private void applyHexFill(CellStyle style, String hexColor) {
        if (!(style instanceof XSSFCellStyle xssfStyle) || hexColor == null || hexColor.isBlank()) {
            return;
        }

        xssfStyle.setFillForegroundColor(new XSSFColor(hexToRgb(hexColor), new DefaultIndexedColorMap()));
        xssfStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private byte[] hexToRgb(String hexColor) {
        String normalized = hexColor.replace("#", "").trim();
        if (normalized.length() != 6) {
            return new byte[]{(byte) 255, (byte) 255, (byte) 255};
        }

        return new byte[]{
                (byte) Integer.parseInt(normalized.substring(0, 2), 16),
                (byte) Integer.parseInt(normalized.substring(2, 4), 16),
                (byte) Integer.parseInt(normalized.substring(4, 6), 16)
        };
    }

    private void writeCell(Row row, int column, String value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private void drawCenteredString(Graphics2D g, String text, int x, int y, int width, int height, FontMetrics metrics) {
        String safeText = text == null ? "" : text;
        int textX = x + Math.max(0, (width - metrics.stringWidth(safeText)) / 2);
        int textY = y + ((height - metrics.getHeight()) / 2) + metrics.getAscent();
        g.drawString(safeText, textX, textY);
    }

    private void drawLeftString(Graphics2D g, String text, int x, int y, int width, int height, FontMetrics metrics) {
        String safeText = text == null ? "" : text;
        int textY = y + ((height - metrics.getHeight()) / 2) + metrics.getAscent();
        g.drawString(safeText, x, textY);
    }

    private void writeNumericCell(Row row, int column, double value, CellStyle style) {
        Cell cell = row.createCell(column);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private void writeSmartNumericCell(
            Row row,
            int column,
            double value,
            CellStyle integerStyle,
            CellStyle decimalStyle
    ) {
        if (Math.abs(value - Math.rint(value)) < 0.0000001d) {
            writeNumericCell(row, column, value, integerStyle);
            return;
        }
        writeNumericCell(row, column, value, decimalStyle);
    }

    private byte[] exportXlsx(JasperPrint print) throws Exception {
        JRXlsxExporter exporter = new JRXlsxExporter();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        exporter.setExporterInput(new SimpleExporterInput(print));
        exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(out));
        exporter.exportReport();
        return out.toByteArray();
    }

    private byte[] exportDocx(JasperPrint print) throws Exception {
        JRDocxExporter exporter = new JRDocxExporter();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        exporter.setExporterInput(new SimpleExporterInput(print));
        exporter.setExporterOutput(new SimpleOutputStreamExporterOutput(out));
        exporter.exportReport();
        return out.toByteArray();
    }

    private byte[] exportHtml(JasperPrint print) throws Exception {
        HtmlExporter exporter = new HtmlExporter();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        exporter.setExporterInput(new SimpleExporterInput(print));
        exporter.setExporterOutput(new SimpleHtmlExporterOutput(out));
        exporter.exportReport();
        return out.toByteArray();
    }

    private int toInt(Object value) {
        if (value == null) {
            return 0;
        }
        if (value instanceof Number number) {
            return number.intValue();
        }
        if (value instanceof String str) {
            String normalized = str.trim();
            if (normalized.isBlank()) {
                return 0;
            }
            try {
                return Integer.parseInt(normalized);
            } catch (NumberFormatException ignored) {
                return 0;
            }
        }
        return 0;
    }
}
