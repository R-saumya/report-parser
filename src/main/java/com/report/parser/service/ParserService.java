package com.report.parser.service;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;

@Service
@RequiredArgsConstructor
public class ParserService {
    private final ObjectMapper objectMapper;
    private final ReportGenerationService reportGenerationService;

    public Resource buildDataAndGetReport() throws IOException {
        ClassPathResource classPathResource = new ClassPathResource("profile.png");
        Resource photoResource = new InputStreamResource(classPathResource.getInputStream());

        ClassPathResource payloadResource = new ClassPathResource("Payload.json");
        Map<String, Object> payload = objectMapper.readValue(payloadResource.getInputStream(),
                new TypeReference<Map<String, Object>>() {
                });
        payload.put("profilePhoto", photoResource);
        Resource template = new ClassPathResource("MyProfile.docx");

        LinkedHashMap<String, String> columnMap = new LinkedHashMap<>();
        columnMap.put("slNo", "S/N");
        columnMap.put("name", "Name");
        columnMap.put("relationship", "Relationship");
        columnMap.put("mobile", "Mobile");
        return reportGenerationService.parseToPdf(payload, template, true, columnMap);
    }
}
