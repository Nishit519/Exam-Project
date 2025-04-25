package com.example.exam.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.example.exam.entity.AdminUser;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class ExcelHelper {

    public static byte[] usersToExcel(List<AdminUser> users) throws IOException {
        String[] columns = {"ID", "Name", "Email", "Date of Birth", "Gender", "Pin Code", "Contact Number", "Address"};
        
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Users Information");

           
            CellStyle cellStyle = workbook.createCellStyle();
            short borderBoundry = 1;

            cellStyle.setBorderTop(borderBoundry);
            cellStyle.setBorderRight(borderBoundry);
            cellStyle.setBorderBottom(borderBoundry);
            cellStyle.setBorderLeft(borderBoundry);

            cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
            cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);

            CellStyle whiteBackground = workbook.createCellStyle();
            whiteBackground.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            whiteBackground.setFillPattern(CellStyle.SOLID_FOREGROUND);
            whiteBackground.setBorderTop(borderBoundry);
            whiteBackground.setBorderRight(borderBoundry);
            whiteBackground.setBorderBottom(borderBoundry);
            whiteBackground.setBorderLeft(borderBoundry);
            whiteBackground.setAlignment(CellStyle.ALIGN_LEFT);
            whiteBackground.setVerticalAlignment(CellStyle.ALIGN_CENTER);

            CellStyle brownBackground = workbook.createCellStyle();
            brownBackground.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());  
            brownBackground.setFillPattern(CellStyle.SOLID_FOREGROUND);
            brownBackground.setBorderTop(borderBoundry);
            brownBackground.setBorderRight(borderBoundry);
            brownBackground.setBorderBottom(borderBoundry);
            brownBackground.setBorderLeft(borderBoundry);
            brownBackground.setAlignment(CellStyle.ALIGN_LEFT);
            brownBackground.setVerticalAlignment(CellStyle.ALIGN_CENTER);

            Row firstRow = sheet.createRow(0);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.length - 1)); // Merging cells from 0 to columns.length - 1 (first row)

            Cell titleCell = firstRow.createCell(0);
            titleCell.setCellValue("User Details");

            CellStyle titleStyle = workbook.createCellStyle();
            titleStyle.setAlignment(CellStyle.ALIGN_CENTER); 
            titleStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER); 
            titleCell.setCellStyle(titleStyle); 

            Row headerRow = sheet.createRow(1);
            for (int col = 0; col < columns.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(columns[col]);
                cell.setCellStyle(cellStyle);  
            }

            int rowIdx = 2; 
            for (AdminUser user : users) {
                Row row = sheet.createRow(rowIdx++);
                
                CellStyle rowStyle = (rowIdx % 2 == 0) ? whiteBackground : brownBackground;

                row.createCell(0).setCellValue(user.getId());
                row.createCell(1).setCellValue(user.getName());
                row.createCell(2).setCellValue(user.getEmail());
                row.createCell(3).setCellValue(user.getDob().toLocalDate().format(DateTimeFormatter.ISO_LOCAL_DATE));
                row.createCell(4).setCellValue(user.getGender().toString());
                row.createCell(5).setCellValue(user.getPinCode());
                row.createCell(6).setCellValue(user.getContactNumber());
                row.createCell(7).setCellValue(user.getAddress());
                
                for (int col = 0; col < columns.length; col++) {
                    row.getCell(col).setCellStyle(rowStyle);
                }
            }
            
            workbook.write(out);
            return out.toByteArray();
        }
    }
}
