package org.example.please1;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;

@RestController
@RequestMapping("/api/tutorials")
public class TutorialController {


    @Autowired
    private TutorialRepository tutorialRepository;

    @PostMapping("/upload")
    public ResponseEntity<String> uploadExcelFile(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("Файл не выбран");
        }

        try (InputStream inputStream = file.getInputStream()) {
            // Используем XSSFWorkbook для работы с .xlsx файлами
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);  // Читаем первый лист

            // Считываем данные из Excel и записываем в базу данных
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {  // Пропускаем заголовок
                    continue;
                }

                // Получаем значения из ячеек
                int id = (int) row.getCell(0).getNumericCellValue();
                String title = row.getCell(1).getStringCellValue();
                String description = row.getCell(2).getStringCellValue();

                // Обрабатываем ячейку с булевым значением
                Cell publishedCell = row.getCell(3);
                boolean published = false;  // значение по умолчанию

                if (publishedCell.getCellType() == CellType.BOOLEAN) {
                    published = publishedCell.getBooleanCellValue();
                } else if (publishedCell.getCellType() == CellType.STRING) {
                    // Преобразуем строку в булево значение
                    String cellValue = publishedCell.getStringCellValue().trim();
                    published = cellValue.equalsIgnoreCase("true") || cellValue.equalsIgnoreCase("yes");
                }

                // Сохраняем данные в базу
                Tutorial tutorial = new Tutorial();
                tutorial.setId((long) id);
                tutorial.setTitle(title);
                tutorial.setDescription(description);
                tutorial.setPublished(published);
                tutorialRepository.save(tutorial);
            }

            return ResponseEntity.status(HttpStatus.OK).body("Данные успешно загружены в базу данных");

        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Ошибка обработки файла");
        }

    }

    @Autowired
    private TutorialService excelService;

    @GetMapping("/download")
    public ResponseEntity<String> downloadExcelFile() throws IOException {
        // Вызываем сервис для сохранения файла на диск
        String filePath = excelService.exportEmployeesToExcel();

        // Возвращаем ответ с подтверждением и путем к файлу
        return ResponseEntity.ok("Excel файл сохранен по пути: " + filePath);
    }

}
