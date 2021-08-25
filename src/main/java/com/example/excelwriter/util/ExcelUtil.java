package com.example.excelwriter.util;

import com.fasterxml.jackson.databind.JavaType;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.Iterator;
import java.util.List;

@Component
@Slf4j
public class ExcelUtil {

    @Autowired
    private ObjectMapper objectMapper;

    public byte[] listToExcel(List<?> data) throws IOException {
        return arrayNodeToExcel(objectMapper.valueToTree(data));
    }

    public byte[] arrayNodeToExcel(ArrayNode arrayNode) throws IOException {
        byte[] excelData = null;
        if (arrayNode.size() > 0) {
            try (SXSSFWorkbook workbook = new SXSSFWorkbook(); ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
                workbook.setCompressTempFiles(true);
                Sheet sheet = workbook.createSheet();
                //header
                Row row = sheet.createRow(0);
                Iterator<String> it1 = arrayNode.get(0).fieldNames();
                int index = 0;
                while(it1.hasNext()) {
                    row.createCell(index++).setCellValue(it1.next());
                }

                //data
                for (int i = 0; i < arrayNode.size(); i++) {
                    JsonNode jsonNode = arrayNode.get(i);
                    Iterator<JsonNode> it2 = jsonNode.iterator();
                    row = sheet.createRow(i + 1);
                    index = 0;
                    while(it2.hasNext()) {
                        setCellValue(row.createCell(index++), it2.next());
                    }
                }

                workbook.write(baos);
                excelData = baos.toByteArray();
                workbook.dispose();
            }
        }
        return excelData;
    }

    public <T> List<T> excelToList(byte[] excelData, Class<T> clazz) throws IOException {
        ArrayNode arrayNode = excelToArrayNode(excelData);
        return objectMapper.readValue(arrayNode.toString(), constructListType(clazz));
    }

    public ArrayNode excelToArrayNode(byte[] excelData) throws IOException {
        ArrayNode arrayNode = objectMapper.createArrayNode();
        try (InputStream is = new ByteArrayInputStream(excelData)) {
            Workbook workbook = new XSSFWorkbook(is);
            Sheet sheet = workbook.getSheetAt(0);
            Row firstRow = sheet.getRow(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                ObjectNode objectNode = objectMapper.createObjectNode();
                Row row = sheet.getRow(i);
                for (int j = 0; j < firstRow.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    getCellValue(cell, objectNode, firstRow.getCell(j).getStringCellValue());
                }
                arrayNode.add(objectNode);
            }
        }
        return arrayNode;
    }

    private void setCellValue(Cell cell, JsonNode cellData) {
        if (cellData.isInt()) {
            cell.setCellValue(cellData.asInt());
        } else if (cellData.isLong()) {
            cell.setCellValue(cellData.asLong());
        } else if (cellData.isDouble()) {
            cell.setCellValue(cellData.asDouble());
        } else if (cellData.isBigDecimal()){
            cell.setCellValue(new BigDecimal(cellData.toString()).doubleValue());
        } else if (cellData.isBoolean()) {
            cell.setCellValue(cellData.asBoolean());
        } else {
            cell.setCellValue(cellData.asText());
        }
    }

    private void getCellValue(Cell cell, ObjectNode objectNode, String fieldName) {
        CellType cellType = cell.getCellType();

        if (CellType.STRING == cellType) {
            if ("null".equals(cell.getStringCellValue())) {
                objectNode.putNull(fieldName);
            } else {
                objectNode.put(fieldName, cell.getStringCellValue());
            }
        } else if (CellType.NUMERIC == cellType) {
            objectNode.put(fieldName, cell.getNumericCellValue());
        } else if (CellType.FORMULA == cellType) {
            objectNode.put(fieldName, cell.getNumericCellValue());
        } else if (CellType.BOOLEAN == cellType) {
            objectNode.put(fieldName, cell.getBooleanCellValue());
        } else {
            log.warn("Unsupported cellType: {}, fieldName: {}", cellType, fieldName);
        }
    }

    private <T> JavaType constructListType(Class<T> clazz) {
        return objectMapper.getTypeFactory().constructCollectionType(List.class, clazz);
    }
}
