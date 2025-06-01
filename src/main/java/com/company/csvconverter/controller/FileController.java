package com.company.csvconverter.controller;

import com.company.csvconverter.service.DataProcessingService;
import com.company.csvconverter.service.ExcelWriterService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
public class FileController {
    
    
    @Autowired
    private DataProcessingService dataProcessingService; // NEW: For 10-column processing
    
    @Autowired
    private ExcelWriterService excelWriterService; // NEW: For dual Excel generation
    
    @GetMapping("/")
    public String uploadPage() {
        return "upload";
    }
    
    /**
     * Original convert endpoint - KEEP for backwards compatibility
     */
    @PostMapping("/convert")
    public ResponseEntity<ByteArrayResource> convertCsvToExcel(@RequestParam("file") MultipartFile file) {
        try {
            // Process CSV file (extract 8 columns + add 2 classifications = 10 columns)
            List<String[]> classifiedData = dataProcessingService.processCsvFile(file);
            
            // Create the final formatted Excel report
            byte[] excelData = excelWriterService.createExcelReport(classifiedData);
            
            ByteArrayResource resource = new ByteArrayResource(excelData);
            
            // Generate unique filename with timestamp
            String filename = "zonal-interchange-report-" + getCurrentTimestamp() + ".xlsx";
            
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .contentLength(excelData.length)
                    .body(resource);
                    
        } catch (Exception e) {
            System.err.println("❌ Error in /convert endpoint: " + e.getMessage());
            return ResponseEntity.badRequest().build();
        }
    }
    
    /**
     * NEW: Upload endpoint for dual downloads (intermediate + final Excel)
     */
    @PostMapping("/upload")
    @ResponseBody
    public ResponseEntity<?> uploadFile(@RequestParam("file") MultipartFile file) {
        try {
            // Process the CSV file (extract 8 columns + add 2 classifications = 10 columns)
            List<String[]> classifiedData = dataProcessingService.processCsvFile(file);
            
            // Create BOTH Excel files
            byte[] intermediateExcel = excelWriterService.createIntermediateExcel(classifiedData);
            byte[] finalExcel = excelWriterService.createExcelReport(classifiedData);
            
            // Create response with both files
            Map<String, Object> response = new HashMap<>();
            response.put("intermediateExcel", Base64.getEncoder().encodeToString(intermediateExcel));
            response.put("finalExcel", Base64.getEncoder().encodeToString(finalExcel));
            response.put("intermediateFileName", "intermediate-10-columns-" + getCurrentTimestamp() + ".xlsx");
            response.put("finalFileName", "zonal-interchange-report-" + getCurrentTimestamp() + ".xlsx");
            response.put("message", "✅ Both Excel files generated successfully!");
            response.put("dataRows", classifiedData.size() - 1); // Exclude header
            
            System.out.println("✅ Dual Excel files generated successfully!");
            System.out.println("📊 Processed " + (classifiedData.size() - 1) + " data rows");
            
            return ResponseEntity.ok(response);
            
        } catch (Exception e) {
            System.err.println("❌ Error in /upload endpoint: " + e.getMessage());
            e.printStackTrace();
            
            Map<String, Object> errorResponse = new HashMap<>();
            errorResponse.put("error", true);
            errorResponse.put("message", "❌ Error processing file: " + e.getMessage());
            
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(errorResponse);
        }
    }
    
    /**
     * Helper method to get current timestamp for unique filenames
     */
    private String getCurrentTimestamp() {
        return LocalDateTime.now()
            .format(DateTimeFormatter.ofPattern("yyyy-MM-dd-HH-mm-ss"));
    }
}