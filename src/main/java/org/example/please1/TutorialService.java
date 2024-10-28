package org.example.please1;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;


import java.io.*;
import java.util.List;

@Service
public class TutorialService {

    @Autowired
    private TutorialRepository employeeRepository;

    public ByteArrayInputStream readEmployeesToExcel() throws IOException {
        List<Tutorial> employees = employeeRepository.findAll();

        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Tutorial");

            // Заголовки
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("ID");
            headerRow.createCell(1).setCellValue("Title");
            headerRow.createCell(2).setCellValue("Description");
            headerRow.createCell(3).setCellValue("Published");

            // Данные
            int rowIdx = 1;
            for (Tutorial employee : employees) {
                Row row = sheet.createRow(rowIdx++);
                row.createCell(0).setCellValue(employee.getId());
                row.createCell(1).setCellValue(employee.getTitle());
                row.createCell(2).setCellValue(employee.getDescription());
                row.createCell(3).setCellValue(employee.isPublished());
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        }
    }

    public String exportEmployeesToExcel() throws IOException {
        List<Tutorial> employees = employeeRepository.findAll();

        // Путь к папке /src/main/resources/
        String filePath = "./src/main/resources/employees.xlsx";
        File file = new File(filePath);

        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOut = new FileOutputStream(file)) {
            Sheet sheet = workbook.createSheet("Employees");

            // Заголовки
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("ID");
            headerRow.createCell(1).setCellValue("Title");
            headerRow.createCell(2).setCellValue("Description");
            headerRow.createCell(3).setCellValue("Published");

            // Данные
            int rowIdx = 1;
            for (Tutorial employee : employees) {
                Row row = sheet.createRow(rowIdx++);
                row.createCell(0).setCellValue(employee.getId());
                row.createCell(1).setCellValue(employee.getTitle());
                row.createCell(2).setCellValue(employee.getDescription());
                row.createCell(3).setCellValue(employee.isPublished());
            }

            // Запись в файл
            workbook.write(fileOut);
        }

        return filePath;  // Возвращаем путь к файлу
    }
}
