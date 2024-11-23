package com.report.parser.service;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.xml.bind.JAXBElement;
import jakarta.xml.bind.JAXBException;
import lombok.RequiredArgsConstructor;
import lombok.extern.log4j.Log4j2;
import org.docx4j.Docx4J;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.springframework.core.io.FileSystemResource;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.springframework.core.io.Resource;

@Service
@Log4j2
@RequiredArgsConstructor
public class ReportGenerationService extends Mapper {
    String regex = "\\$\\{[^{}]+\\}";
    Pattern pattern = Pattern.compile(regex);
    private final ObjectMapper objectMapper;

    public Resource parseToPdf(Map<String, Object> placeholderMap, Resource template, boolean tableInclude, LinkedHashMap<String, String> columnMap) {
        try {
            WordprocessingMLPackage wordMLPackage = prepareWordMLPackage(placeholderMap, template);

            // if there is some table content then add table
            if (tableInclude) {
                List<Map<String, String>> tableData = objectMapper.convertValue(placeholderMap.get("tableData"),
                        new TypeReference<List<Map<String, String>>>() {});

                addTableToDocument(wordMLPackage, tableData, columnMap);
            }
            checkTextAlignment(wordMLPackage);
            // Export to PDF
            wordMLPackage.setFontMapper(new ReportGenerationService(objectMapper));
            OutputStream os = new java.io.FileOutputStream("report.pdf");
            Docx4J.toPDF(wordMLPackage, os);

            return new FileSystemResource("report.pdf");
        } catch (Exception e) {
            log.info("Error occurred: {}", e.getMessage());
        }
        return null;
    }

    /*public Resource parseToWord(Map<String, String> placeholderMap, Resource template) {
        try {
            WordprocessingMLPackage wordMLPackage = prepareWordMLPackage(placeholderMap, template);

            // below line to modify the docx file
            wordMLPackage.save(new java.io.FileOutputStream("quotation.docx"));
            return new FileSystemResource("quotation.docx");
        } catch (Exception e) {
            log.info("Error occurred: {}", e.getMessage());
        }
        return null;
    }*/

    private WordprocessingMLPackage prepareWordMLPackage(Map<String, Object> placeholderMap, Resource template) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(template.getInputStream());
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        // find the defined keys in docx file and verify with placeholderMap
        Set<String> requestedKeys = new HashSet<>();
        getKeysFromDocxFile(documentPart, requestedKeys);
        verifyPlaceholderMap(requestedKeys, placeholderMap);

        VariablePrepare.prepare(wordMLPackage);
        Map<String, String> textPlaceholderMap = new HashMap<>();
        Map<String, Resource> imagePlaceholderMap = new HashMap<>();

        // separate text and image placeholders
        for (Map.Entry<String, Object> entry : placeholderMap.entrySet()) {
            if (entry.getValue() instanceof String) {
                textPlaceholderMap.put(entry.getKey(), (String) entry.getValue());
            } else if (entry.getValue() instanceof Resource) {
                imagePlaceholderMap.put(entry.getKey(), (Resource) entry.getValue());
            }
        }
        replaceTextPlaceholders(documentPart, textPlaceholderMap);
        replaceImagePlaceholders(documentPart, imagePlaceholderMap, wordMLPackage);
        return wordMLPackage;
    }

    private void replaceTextPlaceholders(MainDocumentPart documentPart, Map<String, String> textPlaceholderMap) throws JAXBException, Docx4JException {
        documentPart.variableReplace(textPlaceholderMap);
    }

    private void replaceImagePlaceholders(MainDocumentPart documentPart, Map<String, Resource> imagePlaceholderMap, WordprocessingMLPackage wordMLPackage) throws Exception {
        for (Map.Entry<String, Resource> entry : imagePlaceholderMap.entrySet()) {
            addImageToPlaceholder(documentPart, entry.getKey(), entry.getValue(), wordMLPackage);
        }
    }

    private void verifyPlaceholderMap(Set<String> requestedKeys, Map<String, Object> placeholderMap) {
        List<String> keysInPlaceHolderMap = placeholderMap.keySet().stream().map(key -> String.format("${%s}", key))
                .collect(Collectors.toCollection(ArrayList::new));
        requestedKeys.forEach(key -> {
            if (!keysInPlaceHolderMap.contains(key)) {
                String actualKey = key.substring(key.indexOf("{") + 1, key.indexOf("}"));
                placeholderMap.put(actualKey, "      ");
            }
        });
    }

    private void addImageToPlaceholder(MainDocumentPart documentPart, String placeholder, Resource imageResource, WordprocessingMLPackage wordMLPackage) throws Exception {
        try (InputStream imageInputStream = imageResource.getInputStream()) {
            byte[] imageBytes = imageInputStream.readAllBytes();
            BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, imageBytes);
            Inline inline = imagePart.createImageInline(imageResource.getFilename(), "Image Not Available", 0, 1, false);
            List<Object> elements = documentPart.getJAXBNodesViaXPath("//w:t[text()='" + placeholder + "']", true);
            if (!elements.isEmpty()) {
                for (Object element : elements) {
                    Text placeholderText = extractTextFromElement(element);
                    if (placeholderText != null) {
                        Object parent = placeholderText.getParent();
                        if (parent instanceof R parentRun) {
                            P parentParagraph = (P) parentRun.getParent();
                            Object paragraphParent = parentParagraph.getParent();
                            if (paragraphParent instanceof Tc parentCell) {
                                replaceContentInTableCell(parentCell, placeholderText, inline);
                            } else {
                                replaceContent(documentPart, parentParagraph, inline);
                            }
                        } else if (parent instanceof P) {
                            P parentParagraph = (P) parent;
                            replaceContent(documentPart, parentParagraph, inline);
                        } else if (parent instanceof Tc parentCell) {
                            replaceContentInTableCell(parentCell, placeholderText, inline);
                        }
                    } else {
                        log.warn("Placeholder text not found for key: {}", placeholder);
                    }
                }
            } else {
                log.warn("Image placeholder not found for key: {}", placeholder);
            }
        } catch (Exception e) {
            log.error("Error while replacing image placeholder: ", e);
        }
    }

    private Text extractTextFromElement(Object element) {
        if (element instanceof JAXBElement<?> jaxbElement) {
            if (jaxbElement.getValue() instanceof Text) {
                return (Text) jaxbElement.getValue();
            }
        } else if (element instanceof Text) {
            return (Text) element;
        }
        return null;
    }

    private void replaceContent(MainDocumentPart documentPart, P oldParagraph, Inline imgInline) {
        int index = documentPart.getContent().indexOf(oldParagraph);
        if (index != -1) {
            P imgP = newImageParagraph(imgInline, 200, 100);
            documentPart.getContent().set(index, imgP);
        } else {
            log.warn("Could not find the old paragraph in the document content.");
        }
    }

    private void replaceContentInTableCell(Tc cell, Text placeholderText, Inline imgInline) {
        List<Object> cellContent = cell.getContent();
        for (Object obj : cellContent) {
            if (obj instanceof P paragraph) {
                List<Object> paragraphContent = paragraph.getContent().stream().map(c -> {
                    if (c instanceof R row) {
                        return row.getContent();
                    } else {
                        return null;
                    }
                }).filter(Objects::nonNull).flatMap(List::stream).map(je -> {
                    if (je instanceof JAXBElement) {
                        return ((JAXBElement<?>) je).getValue();
                    } else {
                        return null;
                    }
                }).filter(Objects::nonNull).collect(Collectors.toCollection(ArrayList::new));
                if (paragraphContent.contains(placeholderText)) {
                    P resizedImageParagraph = newImageParagraph(imgInline, getCellWidth(cell) - 10, getCellWidth(cell));
                    int index = cellContent.indexOf(paragraph);
                    cellContent.set(index, resizedImageParagraph);
                    return;
                }
            }
        }
        log.warn("Could not find the placeholder text in the table cell.");
    }

    public P newImageParagraph(Inline inline, long cellWidthPx, long cellHeightPx) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        long emusPerPixel = 9525;
        long cellWidthEmu = cellWidthPx * emusPerPixel;
        long cellHeightEmu = cellHeightPx * emusPerPixel;

        inline.getExtent().setCx(cellWidthEmu);
        inline.getExtent().setCy(cellHeightEmu);

        P p = factory.createP();
        R run = factory.createR();
        p.getContent().add(run);

        Drawing drawing = factory.createDrawing();
        run.getContent().add(drawing);
        drawing.getAnchorOrInline().add(inline);

        return p;
    }

    private long getCellWidth(Tc cell) {
        TcPr tcPr = cell.getTcPr();
        if (tcPr != null && tcPr.getTcW() != null) {
            long value = tcPr.getTcW().getW().longValue();
            return (value * 96) / 1440;
        }
        return 0;
    }

    /*private WordprocessingMLPackage prepareWordMLPackage(Map<String, String> placeholderMap, Resource template) throws Exception {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(template.getInputStream());
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        // find the defined keys in docx file and verify with placeholderMap
        Set<String> requestedKeys = new HashSet<>();
        getKeysFromDocxFile(documentPart, requestedKeys);
        verifyPlaceholderMap(requestedKeys, placeholderMap);

        // Escape XML special characters in the placeholder values
        Map<String, String> escapedPlaceholderMap = new HashMap<>();
        for (Map.Entry<String, String> entry : placeholderMap.entrySet()) {
            escapedPlaceholderMap.put(entry.getKey(), escapeXml(entry.getValue()));
        }

        VariablePrepare.prepare(wordMLPackage);
        documentPart.variableReplace(escapedPlaceholderMap);
        return wordMLPackage;
    }*/

    /*private void verifyPlaceholderMap(Set<String> requestedKeys, Map<String, String> placeholderMap) {
        List<String> keysInPlaceHolderMap = placeholderMap.keySet().stream().map(key -> String.format("${%s}", key)).collect(Collectors.toList());
        requestedKeys.forEach(key -> {
            if (!keysInPlaceHolderMap.contains(key)) {
                String actualKey = key.substring(key.indexOf("{") + 1, key.indexOf("}"));
                placeholderMap.put(actualKey, "      ");
            }
        });
    }*/

    private static void checkTextAlignment(WordprocessingMLPackage wordMLPackage) {
        Body body = wordMLPackage.getMainDocumentPart().getJaxbElement().getBody();
        List<Object> paragraphs = body.getContent();
        for (Object paragraph : paragraphs) {
            if (paragraph instanceof P) {
                P p = (P) paragraph;
                preserveParagraphAlignment(p);
            }
        }
    }

    private static void preserveParagraphAlignment(P p) {
        PPr ppr = p.getPPr();
        if (ppr != null) {
            Jc jc = ppr.getJc();
            if (jc == null || jc.getVal().equals(JcEnumeration.BOTH)) {
                jc = new Jc();
                jc.setVal(JcEnumeration.LEFT);
                ppr.setJc(jc);
            }
        }
    }

    @Override
    public void populateFontMappings(Set<String> set, Fonts fonts) {
        for (String fontName : set) {
            PhysicalFont font = PhysicalFonts.get(fontName);
            if (font == null) {
                font = PhysicalFonts.get("Times New Roman");
            }
            if (font != null) {
                put(fontName, font);
            }
        }
    }

    private void getKeysFromDocxFile(MainDocumentPart documentPart, Set<String> requestedKeys) {
        List<Object> content = documentPart.getContent();
        for (Object obj : content) {
            if (obj instanceof P) {
                addRequestedKeys((P) obj, requestedKeys);
            } else if (obj instanceof JAXBElement<?> && ((JAXBElement<?>) obj).getDeclaredType().equals(org.docx4j.wml.Tbl.class)) {
                processTable((JAXBElement<?>) obj, requestedKeys);
            }
        }
    }

    private void processTable(JAXBElement<?> tableElement, Set<String> requestedKeys) {
        org.docx4j.wml.Tbl tbl = (org.docx4j.wml.Tbl) tableElement.getValue();
        List<Object> rows = tbl.getContent();
        for (Object row : rows) {
            if (row instanceof org.docx4j.wml.Tr) {
                Tr eachRow = (Tr) row;
                List<Object> cellContent = eachRow.getContent();
                for (Object cell : cellContent) {
                    if (cell instanceof JAXBElement<?>) {
                        processCell((JAXBElement<?>) cell, requestedKeys);
                    }
                }
            }
        }
    }

    private void processCell(JAXBElement<?> cellElement, Set<String> requestedKeys) {
        org.docx4j.wml.Tc tc = (org.docx4j.wml.Tc) cellElement.getValue();
        List<Object> cellContent = tc.getContent();
        for (Object cellItem : cellContent) {
            if (cellItem instanceof P) {
                addRequestedKeys((P) cellItem, requestedKeys);
            }
        }
    }

    private void addRequestedKeys(P p, Set<String> requestedKeys) {
        StringBuilder textContent = new StringBuilder();
        List<Object> paragraphContent = p.getContent();
        for (Object item : paragraphContent) {
            if (item instanceof R) {
                R run = (R) item;
                List<Object> runContent = run.getContent();
                for (Object runItem : runContent) {
                    if (runItem instanceof JAXBElement) {
                        JAXBElement<?> element = (JAXBElement<?>) runItem;
                        if (element.getName().getLocalPart().equals("t")) {
                            textContent.append(((Text) element.getValue()).getValue());
                        }
                    }
                }
            }
        }
        String content = textContent.toString();
        Matcher matcher = pattern.matcher(content);
        while (matcher.find()) {
            requestedKeys.add(matcher.group());
        }
    }

    private static void addTableToDocument(WordprocessingMLPackage wordMLPackage, List<Map<String, String>> records, LinkedHashMap<String, String> columnNames) {
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        List<Object> content = documentPart.getContent();

        // Find the table
        for (Object obj : content) {
            if (obj instanceof JAXBElement<?>
                    && ((JAXBElement) obj).getDeclaredType().getName().equalsIgnoreCase("org.docx4j.wml.Tbl")) {
                    Object value = ((JAXBElement<?>) obj).getValue();
                    if (value instanceof Tbl table) {
                        String firstCellValue = getFirstCellValue(table);
                        if (columnNames.containsValue(firstCellValue)) {
                            // Add records as new rows at the top
                            int rowIndex = 1;
                            for (Map<String, String> record : records) {
                                Tr row = createRowFromRecord(columnNames, record);
                                table.getContent().add(rowIndex++, row); // Insert at the top
                            }
                            if (table.getContent().size() > records.size() + 1 && table.getContent().size() > 5) {
                                table.getContent().remove(records.size() + 2);
                            }
                            break; // Once table is found and updated, exit loop
                        }
                    }
                }

        }
    }

    private static String getFirstCellValue(Tbl table) {
        List<Object> tableContent = table.getContent();
        if (!tableContent.isEmpty() && tableContent.get(0) instanceof Tr firstRow) {
            List<Object> firstRowCells = firstRow.getContent();
            if (!firstRowCells.isEmpty() && firstRowCells.get(0) instanceof JAXBElement<?> jaxbElement) {
                if (jaxbElement.getValue() instanceof Tc firstCell) {
                    return getCellValue(firstCell);
                }
            } else if (!firstRowCells.isEmpty() && firstRowCells.get(0) instanceof Tc firstCell) {
                return getCellValue(firstCell);
            }
        }
        return "";
    }

    private static String getCellValue(Tc cell) {
        for (Object obj : cell.getContent()) {
            if (Objects.nonNull(obj)) {
                return obj.toString();
            }
        }
        return "";
    }

    private static Tr createRowFromRecord(LinkedHashMap<String, String> columnNames, Map<String, String> record) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        Tr row = factory.createTr();

        // Create cells for each column
        for (Map.Entry<String, String> entry : columnNames.entrySet()) {
            String columnName = entry.getKey();
            String columnValue = record.get(columnName);

            Tc cell = factory.createTc();
            P paragraph = factory.createP();
            R run = factory.createR();
            Text text = factory.createText();
            text.setValue(columnValue != null ? columnValue : ""); // Set column value, or empty string if null
            run.getContent().add(text);
            paragraph.getContent().add(run);
            cell.getContent().add(paragraph);
            row.getContent().add(cell);
        }

        return row;
    }

    /*private void addTableToDocument(WordprocessingMLPackage wordMLPackage, List<Map<String, String>> records, LinkedHashMap<String, String> columnNames) {
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        Tbl table = createTable(records, columnNames);

        // Find the placeholder and replace it with the table
        List<Object> content = documentPart.getContent();
        for (int i = 0; i < content.size(); i++) {
            Object obj = content.get(i);
            if (obj instanceof P) {
                P p = (P) obj;
                StringBuilder textContent = new StringBuilder();
                List<Object> paragraphContent = p.getContent();
                for (Object item : paragraphContent) {
                    if (item instanceof R) {
                        R run = (R) item;
                        List<Object> runContent = run.getContent();
                        for (Object runItem : runContent) {
                            if (runItem instanceof JAXBElement) {
                                JAXBElement<?> element = (JAXBElement<?>) runItem;
                                if (element.getName().getLocalPart().equals("t")) {
                                    textContent.append(((Text) element.getValue()).getValue());
                                }
                            }
                        }
                    }
                }
                if (textContent.toString().contains("@{tablePlaceholder}")) {
                    // Remove the placeholder paragraph
                    documentPart.getContent().remove(i);

                    // Insert the table at this location
                    documentPart.getContent().add(i, table);
                    break;
                }
            }
        }
    }*/

    /*private void addTableToDocument(WordprocessingMLPackage wordMLPackage, List<Map<String, String>> records, Map<String, String> columnNames) {
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
        Tbl table = createTable(records, columnNames);
        documentPart.addObject(table);
    }*/

    private Tbl createTable(List<Map<String, String>> records, LinkedHashMap<String, String> columnNames) {
        ObjectFactory factory = new ObjectFactory();
        Tbl table = factory.createTbl();

        // Set table borders
        TblPr tblPr = factory.createTblPr();

        TblBorders borders = factory.createTblBorders();
        CTBorder border = new CTBorder();
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(4));
        border.setColor("000000");

        borders.setTop(border);
        borders.setBottom(border);
        borders.setLeft(border);
        borders.setRight(border);
        borders.setInsideH(border);
        borders.setInsideV(border);

        tblPr.setTblBorders(borders);
        table.setTblPr(tblPr);

        // Create table header row
        Tr headerRow = factory.createTr();
        Set<String> keys = columnNames.keySet();
        for (String key : keys) {
            Tc cell = factory.createTc();
            setCellPropertiesAndValue(factory, columnNames.get(key), cell, headerRow);

            // Set the cell background color
            TcPr tcPr = factory.createTcPr();
            CTShd shd = factory.createCTShd();
            shd.setColor("auto");
            shd.setFill("ADD8E6");  // Yellow color in hex
            tcPr.setShd(shd);
            cell.setTcPr(tcPr);
        }
        table.getContent().add(headerRow);

        // Create table rows
        for (Map<String, String> record : records) {
            Tr row = factory.createTr();
            for (String key : keys) {
                Tc cell = factory.createTc();
                setCellPropertiesAndValue(factory, record.get(key), cell, row);
            }
            table.getContent().add(row);
        }

        return table;
    }

    private static void setCellPropertiesAndValue(ObjectFactory factory, String columnNames, Tc cell, Tr row) {
        // Create paragraph
        P paragraph = factory.createP();

        // Set paragraph alignment to center
        PPr paragraphProperties = factory.createPPr();
        Jc justification = factory.createJc();
        justification.setVal(JcEnumeration.CENTER);
        paragraphProperties.setJc(justification);
        paragraph.setPPr(paragraphProperties);

        // Create run and set text
        R run = factory.createR();
        Text text = factory.createText();
        text.setValue(columnNames);
        run.getContent().add(text);
        paragraph.getContent().add(run);
        cell.getContent().add(paragraph);
        row.getContent().add(cell);

        // Set the font properties
        RPr runProperties = factory.createRPr();

        // Set font to Times New Roman
        RFonts rFonts = factory.createRFonts();
        rFonts.setAscii("Times New Roman");
        rFonts.setHAnsi("Times New Roman");
        runProperties.setRFonts(rFonts);

        // Set font size to 11
        HpsMeasure size = factory.createHpsMeasure();
        size.setVal(BigInteger.valueOf(22)); // 11 * 2 because the unit is half-points
        runProperties.setSz(size);
        runProperties.setSzCs(size);

        // Apply the run properties to the run
        run.setRPr(runProperties);
    }

    private String escapeXml(String input) {
        if (input == null) {
            return null;
        }
        StringBuilder escaped = new StringBuilder();
        for (char c : input.toCharArray()) {
            switch (c) {
                case '&':
                    escaped.append("&amp;");
                    break;
                case '<':
                    escaped.append("&lt;");
                    break;
                case '>':
                    escaped.append("&gt;");
                    break;
                case '"':
                    escaped.append("&quot;");
                    break;
                case '\'':
                    escaped.append("&apos;");
                    break;
                default:
                    escaped.append(c);
                    break;
            }
        }
        return escaped.toString();
    }
}