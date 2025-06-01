package com.company.csvconverter.service;

import com.opencsv.CSVReader;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStreamReader;
import java.util.*;
import java.util.stream.Collectors;
import java.util.HashMap;
import java.util.Map;
import java.util.Arrays;

@Service
public class DataProcessingService {
    
    @Autowired
    private ClassificationService classificationService;
    
    // UPDATED: Define the 8 columns we want to extract from CSV (added ZONE TO)
    private static final String[] TARGET_COLUMNS = {
        "ZONE TO",              // NEW: Added ZONE TO as first column
        "IC STTN",
        "HANDED OVER STTN TO", 
        "HANDED OVER L/E", 
        "HANDED OVER TYPE",
        "TAKEN OVER STTN TO",
        "TAKEN OVER L/E", 
        "TAKEN OVER TYPE"
    };
    
    // UPDATED: Final columns including classification columns (now 10 total)
    private static final String[] FINAL_COLUMNS = {
        "ZONE TO",              // NEW: Added ZONE TO as first column
        "IC STTN",
        "HANDED OVER STTN TO", 
        "HANDED OVER L/E", 
        "HANDED OVER TYPE",
        "HANDED OVER TYPE CLASSIFICATION",
        "TAKEN OVER STTN TO",
        "TAKEN OVER L/E", 
        "TAKEN OVER TYPE",
        "TAKEN OVER TYPE CLASSIFICATION"
    };
    
    /**
     * Main method to process CSV file and return classified data
     */
    public List<String[]> processCsvFile(MultipartFile csvFile) throws Exception {
        // Read CSV data
        List<String[]> csvData = readCsvFile(csvFile);
        
        // Extract required columns starting from row 3
        List<String[]> extractedData = extractRequiredColumns(csvData);
        
        // Add classification columns
        List<String[]> classifiedData = addClassificationColumns(extractedData);
        
        // NEW: Sort the data by ZONE TO order BEFORE returning
        List<String[]> sortedData = sortDataByZone(classifiedData);
        
        return sortedData;
    }
    
    /**
     * Reads CSV file and returns all data as List of String arrays
     */
    private List<String[]> readCsvFile(MultipartFile file) throws Exception {
        try (CSVReader reader = new CSVReader(new InputStreamReader(file.getInputStream()))) {
            List<String[]> data = reader.readAll();
            System.out.println("CSV file read successfully. Total rows: " + data.size());
            return data;
        }
    }
    
    /**
     * Extracts only the required columns from CSV data
     * Starts from row 3 (index 2) as header row
     * Data extraction starts from row 4 (index 3)
     */
    private List<String[]> extractRequiredColumns(List<String[]> csvData) {
        List<String[]> extractedData = new ArrayList<>();
        
        if (csvData.size() < 3) {
            System.out.println("Warning: CSV has less than 3 rows. Using default headers.");
            extractedData.add(TARGET_COLUMNS);
            return extractedData;
        }
        
        // Get the header row (assuming it's row 3, which is index 2)
        String[] headerRow = csvData.get(2);
        System.out.println("Header row found: " + Arrays.toString(headerRow));
        
        // Find column indices for our target columns
        int[] columnIndices = findColumnIndices(headerRow, TARGET_COLUMNS);
        
        // Add our custom headers as first row
        extractedData.add(TARGET_COLUMNS);
        
        // Extract data from row 4 onwards (index 3 onwards)
        int dataRowsProcessed = 0;
        for (int i = 3; i < csvData.size(); i++) {
            String[] currentRow = csvData.get(i);
            String[] extractedRow = new String[TARGET_COLUMNS.length];
            
            // Extract only the required columns
            for (int j = 0; j < columnIndices.length; j++) {
                if (columnIndices[j] != -1 && columnIndices[j] < currentRow.length) {
                    extractedRow[j] = currentRow[columnIndices[j]];
                } else {
                    extractedRow[j] = ""; // Empty if column not found
                }
            }
            
            extractedData.add(extractedRow);
            dataRowsProcessed++;
        }
        
        System.out.println("Data extraction completed. Processed " + dataRowsProcessed + " data rows.");
        return extractedData;
    }
    
    /**
     * UPDATED: Adds classification columns based on wagon types
     * Converts 8-column data to 10-column data with classifications
     */
    private List<String[]> addClassificationColumns(List<String[]> extractedData) {
        List<String[]> classifiedData = new ArrayList<>();
        
        if (extractedData.isEmpty()) {
            System.out.println("Warning: No extracted data to classify.");
            return classifiedData;
        }
        
        // Add the new headers with classification columns
        classifiedData.add(FINAL_COLUMNS);
        
        // Process data rows (skip header row)
        int classifiedRows = 0;
        for (int i = 1; i < extractedData.size(); i++) {
            String[] originalRow = extractedData.get(i);
            String[] newRow = new String[FINAL_COLUMNS.length];
            
            // UPDATED: Copy original data and add classifications (now with ZONE TO)
            newRow[0] = originalRow[0]; // ZONE TO (NEW)
            newRow[1] = originalRow[1]; // IC STTN
            newRow[2] = originalRow[2]; // HANDED OVER STTN TO
            newRow[3] = originalRow[3]; // HANDED OVER L/E
            newRow[4] = originalRow[4]; // HANDED OVER TYPE
            newRow[5] = classifyWagonType(originalRow[4]); // HANDED OVER TYPE CLASSIFICATION
            newRow[6] = originalRow[5]; // TAKEN OVER STTN TO
            newRow[7] = originalRow[6]; // TAKEN OVER L/E
            newRow[8] = originalRow[7]; // TAKEN OVER TYPE
            newRow[9] = classifyWagonType(originalRow[7]); // TAKEN OVER TYPE CLASSIFICATION
            
            classifiedData.add(newRow);
            classifiedRows++;
        }
        
        System.out.println("Classification completed. Classified " + classifiedRows + " rows.");
        return classifiedData;
    }
    
    /**
     * Classifies wagon type using ClassificationService
     */
    private String classifyWagonType(String wagonType) {
        if (wagonType == null || wagonType.trim().isEmpty()) {
            return "";
        }
        
        String classification = classificationService.getClassification(wagonType);
        if (classification == null) {
            System.out.println("No classification found for wagon type: " + wagonType);
            return "";
        }
        
        return classification;
    }
    
    /**
     * Finds the column indices for target columns in the header row
     */
    private int[] findColumnIndices(String[] headerRow, String[] targetColumns) {
        int[] indices = new int[targetColumns.length];
        Arrays.fill(indices, -1); // Initialize with -1 (not found)
        
        for (int i = 0; i < targetColumns.length; i++) {
            for (int j = 0; j < headerRow.length; j++) {
                if (headerRow[j] != null && 
                    headerRow[j].trim().equalsIgnoreCase(targetColumns[i].trim())) {
                    indices[i] = j;
                    break;
                }
            }
        }
        
        // Debug: Print column mapping
        System.out.println("Column Mapping:");
        for (int i = 0; i < targetColumns.length; i++) {
            System.out.println(targetColumns[i] + " -> " + 
                (indices[i] != -1 ? "Column " + indices[i] : "NOT FOUND"));
        }
        
        return indices;
    }
    
    /**
     * Gets the final column structure for other services
     */
    public String[] getFinalColumns() {
        return FINAL_COLUMNS.clone();
    }
    
    /**
     * Gets the target columns for reference
     */
    public String[] getTargetColumns() {
        return TARGET_COLUMNS.clone();
    }
    
    /**
     * Sorts the classified data by ZONE TO in the specified order: CR, WC, NW, DFCR
     * Then sorts IC STTN within each zone in the specified order
     */
    private List<String[]> sortDataByZone(List<String[]> classifiedData) {
        if (classifiedData.isEmpty()) {
            return classifiedData;
        }
        
        // Define the zone order
        String[] zoneOrder = {"CR", "WC", "NW", "DFCR"};
        Map<String, Integer> zoneOrderMap = new HashMap<>();
        for (int i = 0; i < zoneOrder.length; i++) {
            zoneOrderMap.put(zoneOrder[i], i);
        }
        
        // Define IC STTN order within each zone
        Map<String, String[]> icSttnOrderByZone = new HashMap<>();
        icSttnOrderByZone.put("CR", new String[]{"BSR", "JL", "KNW"});
        icSttnOrderByZone.put("WC", new String[]{"SHRN", "NAD", "MKC", "MTA", "CNA"});
        icSttnOrderByZone.put("NW", new String[]{"BEC", "AII", "HMT", "BLDI", "PNU"});
        icSttnOrderByZone.put("DFCR", new String[]{"BHU", "CECC", "GGM", "MSH", "SAU", "MPR", "GTX", "NOL", "SJN", "SAH"});
        
        // Create order maps for IC STTN within each zone
        Map<String, Map<String, Integer>> icSttnOrderMaps = new HashMap<>();
        for (Map.Entry<String, String[]> entry : icSttnOrderByZone.entrySet()) {
            String zone = entry.getKey();
            String[] icSttnOrder = entry.getValue();
            Map<String, Integer> orderMap = new HashMap<>();
            for (int i = 0; i < icSttnOrder.length; i++) {
                orderMap.put(icSttnOrder[i], i);
            }
            icSttnOrderMaps.put(zone, orderMap);
        }
        
        // Separate header and data
        List<String[]> header = new ArrayList<>();
        header.add(classifiedData.get(0));
        List<String[]> dataRows = classifiedData.subList(1, classifiedData.size());
        
        // FIRST: Apply zone transformation to each row BEFORE sorting
        List<String[]> transformedDataRows = dataRows.stream()
            .map(row -> {
                String[] transformedRow = row.clone(); // Clone to avoid modifying original
                String zoneTo = transformedRow[0];
                String icSttn = transformedRow[1];
                
                // Apply zone transformation: NW + CNA → AII
                String transformedIcSttn = applyZoneTransformation(zoneTo, icSttn);
                transformedRow[1] = transformedIcSttn; // Update IC STTN in the row
                
                return transformedRow;
            })
            .collect(Collectors.toList());
        
        // THEN: Sort transformed data rows by ZONE TO, then by IC STTN within zone
        List<String[]> sortedDataRows = transformedDataRows.stream()
            .sorted((row1, row2) -> {
                String zone1 = row1[0] != null ? row1[0].trim().toUpperCase() : "";
                String zone2 = row2[0] != null ? row2[0].trim().toUpperCase() : "";
                String icSttn1 = row1[1] != null ? row1[1].trim().toUpperCase() : "";
                String icSttn2 = row2[1] != null ? row2[1].trim().toUpperCase() : "";
                
                // First sort by zone order
                int zone1Order = zoneOrderMap.getOrDefault(zone1, 999);
                int zone2Order = zoneOrderMap.getOrDefault(zone2, 999);
                
                if (zone1Order != zone2Order) {
                    return Integer.compare(zone1Order, zone2Order);
                }
                
                // If zones are the same, sort by IC STTN order within that zone
                Map<String, Integer> icSttnOrderMap = icSttnOrderMaps.get(zone1);
                if (icSttnOrderMap != null) {
                    int icSttn1Order = icSttnOrderMap.getOrDefault(icSttn1, 999); // 999 for unknown IC STTN
                    int icSttn2Order = icSttnOrderMap.getOrDefault(icSttn2, 999);
                    
                    if (icSttn1Order != icSttn2Order) {
                        return Integer.compare(icSttn1Order, icSttn2Order);
                    }
                    
                    // If both are unknown (order = 999), sort alphabetically
                    if (icSttn1Order == 999 && icSttn2Order == 999) {
                        return icSttn1.compareTo(icSttn2);
                    }
                }
                
                // Fallback: sort alphabetically by IC STTN
                return icSttn1.compareTo(icSttn2);
            })
            .collect(Collectors.toList());
        
        // Combine header and sorted data
        List<String[]> sortedData = new ArrayList<>();
        sortedData.addAll(header);
        sortedData.addAll(sortedDataRows);
        
        System.out.println("✅ Data sorted by ZONE TO (CR, WC, NW, DFCR) then by IC STTN in specified order");
        System.out.println("🔄 Zone transformations applied BEFORE sorting (NW+CNA → AII)");
        return sortedData;
    }
    
    /**
     * Applies zone transformation rules - SAME AS IN ExcelWriterService
     * Rule: If ZONE TO = "NW" and IC STTN = "CNA", change IC STTN to "AII"
     */
    private String applyZoneTransformation(String zoneTo, String icSttn) {
        if (zoneTo == null || icSttn == null) {
            return icSttn; // Return original if either is null
        }
        
        String cleanZoneTo = zoneTo.trim().toUpperCase();
        String cleanIcSttn = icSttn.trim().toUpperCase();
        
        // Apply transformation rule: NW + CNA → AII
        if ("NW".equals(cleanZoneTo) && "CNA".equals(cleanIcSttn)) {
            System.out.println("Zone transformation applied: ZONE TO=" + zoneTo + ", IC STTN changed from " + icSttn + " to AII");
            return "AII";
        }
        
        // Add more transformation rules here if needed in the future
        return icSttn; // Return original IC STTN if no transformation rule applies
    }
}
