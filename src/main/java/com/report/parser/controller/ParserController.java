package com.report.parser.controller;

import com.report.parser.service.ParserService;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@RequiredArgsConstructor
public class ParserController {
    private final ParserService parserService;

    @GetMapping("/generate")
    public ResponseEntity<Resource> getPreviewReportByTableNameAndRecordId() throws IOException {
        Resource resp = parserService.buildDataAndGetReport();
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header(HttpHeaders.CONTENT_DISPOSITION, String.format("attachment; filename=\"%s\"", resp.getFilename()))
                .body(resp);
    }
}
