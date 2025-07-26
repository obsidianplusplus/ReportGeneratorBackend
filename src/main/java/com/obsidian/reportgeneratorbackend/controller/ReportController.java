package com.obsidian.reportgeneratorbackend.controller;

import com.obsidian.reportgeneratorbackend.dto.ExcelChartRequest;
import com.obsidian.reportgeneratorbackend.dto.ReportGenerationRequest;
import com.obsidian.reportgeneratorbackend.service.ReportGenerationService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.Date;

@RestController
@RequestMapping("/api")
@CrossOrigin(origins = "*", exposedHeaders = {"Content-Disposition"})
public class ReportController {

    private final ReportGenerationService reportService;

    public ReportController(ReportGenerationService reportService) {
        this.reportService = reportService;
    }

    @PostMapping(value = "/reports/generate", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> generateReport(
            @RequestPart("template") MultipartFile templateFile,
            @RequestPart("request") ReportGenerationRequest request) {

        try {
            byte[] reportBytes = reportService.generateReport(request, templateFile.getBytes());
            HttpHeaders headers = new HttpHeaders();
            String filename = generateFilename(request);
            headers.setContentDispositionFormData("attachment", URLEncoder.encode(filename, StandardCharsets.UTF_8.name()));

            if (request.getExportMode() == com.obsidian.reportgeneratorbackend.model.ExportMode.ZIP_FILES) {
                headers.setContentType(MediaType.valueOf("application/zip"));
            } else {
                headers.setContentType(MediaType.valueOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            }

            return new ResponseEntity<>(reportBytes, headers, HttpStatus.OK);

        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(e.getMessage().getBytes(StandardCharsets.UTF_8), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    /**
     * 新增：定义从Excel内部提取数据生成图表的接口。
     */
    @PostMapping(value = "/charts/generate-from-excel", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> generateChartFromExcel(
            @RequestPart("template") MultipartFile templateFile,
            @RequestPart("request") ExcelChartRequest request) {

        try {
            byte[] reportBytes = reportService.generateChartInExcel(templateFile.getBytes(), request);
            HttpHeaders headers = new HttpHeaders();
            String filename = "Visualized_Report_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xlsx";

            headers.setContentDispositionFormData("attachment", URLEncoder.encode(filename, StandardCharsets.UTF_8.name()));
            headers.setContentType(MediaType.valueOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));

            return new ResponseEntity<>(reportBytes, headers, HttpStatus.OK);

        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(e.getMessage().getBytes(StandardCharsets.UTF_8), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    private String generateFilename(ReportGenerationRequest request) {
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String baseName = "Generated_Report";
        String extension = ".xlsx";

        switch (request.getExportMode()) {
            case SINGLE_SHEET:
                baseName = "Report_Single_Sheet";
                break;
            case MULTI_SHEET:
                baseName = "Report_Multi_Sheet";
                break;
            case ZIP_FILES:
                baseName = "Report_Archive";
                extension = ".zip";
                break;
        }
        return baseName + "_" + timestamp + extension;
    }
}