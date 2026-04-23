package com.jasper.jasper_service;

import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.data.JRMapCollectionDataSource;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@SpringBootTest
class JasperServiceApplicationTests {

	@Test
	void contextLoads() {
	}

	@Test
	void workPerformedReportCompilesAndFills() throws Exception {
		InputStream jrxml = new ClassPathResource("reports/WorkPerformedReport.jrxml").getInputStream();
		JasperReport report = JasperCompileManager.compileReport(jrxml);

		Map<String, Object> header = new HashMap<>();
		header.put("id", 16);
		header.put("created_at", "2026-04-06 15:25:02.101614");
		header.put("project_name", "Фламинго");
		header.put("block_name", "Блок А");
		header.put("status_name", "Создан");
		header.put("performed_person_name", "ИП Иванов");
		header.put("foreman_name", null);

		Map<String, Object> item = new HashMap<>();
		item.put("service_name", "Монтаж фасада");
		item.put("stage_name", "Тест");
		item.put("subsection_name", "Подэтапы 1");
		item.put("unit_name", "пог. метр");
		item.put("quantity", 1000.0);
		item.put("price", 10000.0);
		item.put("total", 10000000.0);

		Map<String, Object> params = new HashMap<>();
		params.put("total", 10000000.0);
		params.put("totalInWords", "Десять миллионов сом 00 тыйын");
		params.put("reportId", String.valueOf(header.get("id")));
		params.put("createdAt", String.valueOf(header.get("created_at")));
		params.put("projectName", String.valueOf(header.get("project_name")));
		params.put("performerName", String.valueOf(header.get("performed_person_name")));
		params.put("foremanName", "");
		params.put("planningEngineerName", "");
		params.put("mainEngineerName", "");
		params.put("performerSignatureName", "ИП Иванов");
		params.put("foremanSignatureName", "");
		params.put("planningEngineerSignatureName", "");
		params.put("mainEngineerSignatureName", "");

		JasperFillManager.fillReport(
				report,
				params,
				new JRMapCollectionDataSource(List.of(item))
		);
	}

}
