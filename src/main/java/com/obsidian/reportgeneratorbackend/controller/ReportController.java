package com.obsidian.reportgeneratorbackend.controller;

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

/*
 * 描述: API 控制器，定义了报告生成的端点(endpoint)。
 *       它负责接收HTTP请求，调用服务处理，并返回文件响应。
 */
@RestController
@RequestMapping("/api/reports") // 所有请求都以 /api/reports 为前缀
@CrossOrigin(origins = "*") // 允许所有来源的跨域请求，在生产环境中应配置为前端的实际地址
public class ReportController {

    private final ReportGenerationService reportService;

    // 使用构造函数注入服务，这是Spring推荐的方式
    public ReportController(ReportGenerationService reportService) {
        this.reportService = reportService;
    }

    /*
     * 定义报告生成的POST接口。
     * 使用 @RequestPart 来分别接收文件和JSON数据。
     * @param templateFile  上传的Excel模板文件
     * @param request       包含映射规则和日志数据的JSON对象
     * @return 返回一个包含文件内容的HTTP响应
     */
    @PostMapping(value = "/generate", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> generateReport(
            @RequestPart("template") MultipartFile templateFile,
            @RequestPart("request") ReportGenerationRequest request) {

        try {
            // 调用服务层生成报告
            byte[] reportBytes = reportService.generateReport(request, templateFile.getBytes());

            // 准备HTTP响应头
            HttpHeaders headers = new HttpHeaders();
            String filename = generateFilename(request);

            // 设置响应头，告知浏览器这是一个文件下载
            headers.setContentDispositionFormData("attachment", URLEncoder.encode(filename, StandardCharsets.UTF_8));

            // 根据导出模式设置不同的MIME类型
            if (request.getExportMode() == com.obsidian.reportgeneratorbackend.model.ExportMode.ZIP_FILES) {
                headers.setContentType(MediaType.valueOf("application/zip"));
            } else {
                headers.setContentType(MediaType.valueOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            }

            return new ResponseEntity<>(reportBytes, headers, HttpStatus.OK);

        } catch (IOException e) {
            // 处理文件读写错误
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        } catch (Exception e) {
            // 处理其他未知错误
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.BAD_REQUEST);
        }
    }

    /*
     * 根据请求动态生成一个友好的文件名。
     */
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