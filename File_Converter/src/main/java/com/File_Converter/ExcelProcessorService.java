package com.File_Converter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

@Service
public class ExcelProcessorService {

    public void processExcelFile(String filePath, String outputDirectory) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                processSheet(sheet, outputDirectory);
            }
        }
    }

    private void processSheet(Sheet sheet, String outputDirectory) throws IOException {
        int chunkSize = 50000;
        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int startRow = 0; startRow < rowCount; startRow += chunkSize) {
            int endRow = Math.min(startRow + chunkSize, rowCount);
            processChunk(sheet, startRow, endRow, outputDirectory);
        }
    }

    private void processChunk(Sheet sheet, int startRow, int endRow, String outputDirectory) throws IOException {
        try (Workbook chunkWorkbook = new XSSFWorkbook(); // Use XSSFWorkbook for .xlsx format, or HSSFWorkbook for .xls
             FileOutputStream fos = new FileOutputStream(outputDirectory + "/output_chunk_" + startRow + ".xlsx")) {

            Sheet chunkSheet = chunkWorkbook.createSheet("ChunkSheet");
            for (int i = startRow; i < endRow; i++) {
                Row sourceRow = sheet.getRow(i);
                Row destinationRow = chunkSheet.createRow(i - startRow);

                for (int j = 0; j < sourceRow.getPhysicalNumberOfCells(); j++) {
                    Cell sourceCell = sourceRow.getCell(j);
                    Cell destinationCell = destinationRow.createCell(j);

                    // Copy cell value and style
                    if (sourceCell != null) {
                        destinationCell.setCellValue(sourceCell.getStringCellValue()); // You may need to adjust based on your data types
                        destinationCell.setCellStyle(sourceCell.getCellStyle());
                    }
                }
            }

            chunkWorkbook.write(fos);
        }
    }

}
