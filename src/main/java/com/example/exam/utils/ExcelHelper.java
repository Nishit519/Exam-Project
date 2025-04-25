package com.example.exam.utils;

import org.apache.poi.ss.usermodel.*;
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

            
            Row headerRow = sheet.createRow(0);
            for (int col = 0; col < columns.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(columns[col]);
            }

            
            int rowIdx = 1;
            for (AdminUser user : users) {
                Row row = sheet.createRow(rowIdx++);
                row.createCell(0).setCellValue(user.getId());
                row.createCell(1).setCellValue(user.getName());
                row.createCell(2).setCellValue(user.getEmail());
                row.createCell(3).setCellValue(user.getDob().toLocalDate().format(DateTimeFormatter.ISO_LOCAL_DATE));
                row.createCell(4).setCellValue(user.getGender().toString());
                row.createCell(5).setCellValue(user.getPinCode());
                row.createCell(6).setCellValue(user.getContactNumber());
                row.createCell(7).setCellValue(user.getAddress());
            }

            workbook.write(out);
            return out.toByteArray();
        }
    }
}
