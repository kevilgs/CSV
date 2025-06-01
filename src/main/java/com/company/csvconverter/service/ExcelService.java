package com.company.csvconverter.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@Service
public class ExcelService {
    
    @Autowired
    private DataProcessingService dataProcessingService;
    
    @Autowired
    private ExcelWriterService excelWriterService;
    
    /**
     * Main method to convert CSV to Excel
     * This is the only public method - clean and simple!
     */
    public byte[] convertCsvToExcel(MultipartFile csvFile) throws Exception {
        System.out.println("Starting CSV to Excel conversion process...");
        
        // Step 1: Process the CSV file and get classified data
        List<String[]> classifiedData = dataProcessingService.processCsvFile(csvFile);
        
        // Step 2: Create Excel report with the classified data
        byte[] excelReport = excelWriterService.createExcelReport(classifiedData);
        
        System.out.println("CSV to Excel conversion completed successfully!");
        return excelReport;
    }
}