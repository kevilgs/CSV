package com.company.csvconverter.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelWriterService {
    
    @Autowired
    private ExcelStyleService styleService;
    
    @Autowired
    private ExcelStructureService structureService;
    
    /**
     * Creates the complete Excel workbook with classified data
     * Returns byte array ready for download
     */
    public byte[] createExcelReport(List<String[]> classifiedData) throws Exception {
        // Create Excel workbook
        Workbook workbook = new XSSFWorkbook();
        
        // Create the main report sheet
        Sheet reportSheet = workbook.createSheet("Zonal Interchange Report");
        
        // Create the complete report structure (Headers, Styling, Column Widths)
        structureService.createReportStructure(reportSheet, workbook);
        
        // Write the classified data to the sheet
        writeClassifiedDataToSheet(reportSheet, classifiedData, workbook);
        
        // Convert workbook to byte array
        byte[] excelData = convertWorkbookToByteArray(workbook);
        
        // Close workbook to free memory
        workbook.close();
        
        System.out.println("Excel report created successfully with " + (classifiedData.size() - 1) + " data rows.");
        return excelData;
    }
    
    /**
     * Writes the classified data starting from Row 5 (after 4 header rows)
     * CLEAN VERSION: No console logging, just Excel generation
     */
    // Update this part in writeClassifiedDataToSheet method

/**
 * Writes the classified data starting from Row 5 (after 4 header rows)
 * UPDATED: Now handles zone-ordered data
 */
private void writeClassifiedDataToSheet(Sheet sheet, List<String[]> classifiedData, Workbook workbook) {
    if (classifiedData.isEmpty()) {
        System.out.println("Warning: No classified data to write to Excel.");
        return;
    }
    
    // Create simple data style (no bold, no background)
    CellStyle dataStyle = styleService.createDataStyle(workbook);
    
    // Group data by ZONE TO then IC STTN
    Map<String, StationData> groupedData = groupDataByIcSttn(classifiedData);
    
    // Start writing data from Row 5 (index 4) - after our 4 header rows
    int currentRow = 4;
    int dataRowsWritten = 0;
    
    // Write data for each unique IC STTN (now ordered by zone)
    for (Map.Entry<String, StationData> entry : groupedData.entrySet()) {
       // "CR_BSR", "WCR_CNA", etc.
        String icSttn = entry.getValue().icSttn; // Extract actual IC STTN name
        StationData stationData = entry.getValue();
        
        // Write BOTH summary AND detailed station rows starting from the SAME row
        currentRow = writeCompleteStationData(sheet, currentRow, icSttn, stationData, dataStyle);
        dataRowsWritten++;
    }
    
    // Add TOTAL row after all data (with one empty row gap)
    addTotalRow(sheet, currentRow + 1, groupedData, dataStyle); // +1 for empty row gap
    
    System.out.println("✅ Excel report generated successfully!");
    System.out.println("📊 " + dataRowsWritten + " IC STTN stations processed (ordered by CR, WCR, NWR, DFCR)");
    System.out.println("📈 " + (classifiedData.size() - 1) + " total data rows processed");
    System.out.println("🔄 Zone transformations and classifications applied");
    System.out.println("📋 TOTAL row added with calculated sums");
}
    
    /**
     * Writes complete station data with THICK OUTSIDE BORDERS for each IC STTN block
     */
    private int writeCompleteStationData(Sheet sheet, int startRow, String icSttn, StationData data, CellStyle dataStyle) {
        // Find the maximum number of stations across ALL classifications (both sections)
        int maxStations = Math.max(1, calculateMaxStations(data)); // At least 1 for summary row
        
        // Create enough rows for all stations (including summary row)
        for (int i = 0; i < maxStations; i++) {
            Row dataRow = sheet.createRow(startRow + i);
            
            // Fill all columns with empty cells first (only used columns now: A-Y, 25 total)
            for (int col = 0; col <= 24; col++) {
                writeCell(dataRow, col, "", dataStyle);
            }
        }
        
        // Write summary data in the FIRST row (startRow + 0)
        Row summaryRow = sheet.getRow(startRow);
        writeSummaryRowData(summaryRow, icSttn, data, dataStyle);
        
        // Fill HANDEDOVER classification details starting from the SAME row (startRow + 0)
        fillClassificationColumn(sheet, startRow, data.handedOverJumboStations, 4, dataStyle);     // Column E - JUMBO
        fillClassificationColumn(sheet, startRow, data.handedOverBoxnStations, 5, dataStyle);      // Column F - BOXN
        fillClassificationColumn(sheet, startRow, data.handedOverBtpnStations, 6, dataStyle);      // Column G - BTPN
        fillClassificationColumn(sheet, startRow, data.handedOverBtpgStations, 7, dataStyle);      // Column H - BTPG
        fillClassificationColumn(sheet, startRow, data.handedOverContStations, 8, dataStyle);      // Column I - CONT
        fillClassificationColumn(sheet, startRow, data.handedOverShraStations, 9, dataStyle);      // Column J - SHRA
        fillClassificationColumn(sheet, startRow, data.handedOverOthersStations, 10, dataStyle);   // Column K - OTHERS
        fillClassificationColumn(sheet, startRow, data.handedOverEmptiesStations, 11, dataStyle);  // Column L - EMPTIES
        
        // Fill TAKENOVER classification details starting from the SAME row (startRow + 0) - SHIFTED LEFT
        fillClassificationColumn(sheet, startRow, data.takenOverJumboStations, 17, dataStyle);     // Column R - JUMBO (SHIFTED)
        fillClassificationColumn(sheet, startRow, data.takenOverBoxnStations, 18, dataStyle);      // Column S - BOXN (SHIFTED)
        fillClassificationColumn(sheet, startRow, data.takenOverBtpnStations, 19, dataStyle);      // Column T - BTPN (SHIFTED)
        fillClassificationColumn(sheet, startRow, data.takenOverBtpgStations, 20, dataStyle);      // Column U - BTPG (SHIFTED)
        fillClassificationColumn(sheet, startRow, data.takenOverContStations, 21, dataStyle);      // Column V - CONT (SHIFTED)
        fillClassificationColumn(sheet, startRow, data.takenOverShraStations, 22, dataStyle);      // Column W - SHRA (SHIFTED)
        fillClassificationColumn(sheet, startRow, data.takenOverOthersStations, 23, dataStyle);    // Column X - OTHERS (SHIFTED)
        fillClassificationColumn(sheet, startRow, data.takenOverEmptiesStations, 24, dataStyle);   // Column Y - EMPTIES (SHIFTED)

        // Merge cells for this IC STTN block
        mergeCellsForIcSttn(sheet, startRow, maxStations, icSttn, data);

        // NEW: Apply thick outside border around the entire IC STTN block
        applyThickBorderAroundBlock(sheet, startRow, maxStations);

        return startRow + maxStations; // Return the next available row for the next IC STTN
    }

    /**
     * Applies a thick outside border around a block of rows (IC STTN block).
     */
    private void applyThickBorderAroundBlock(Sheet sheet, int startRow, int rowCount) {
        if (rowCount <= 0) return;
        int endRow = startRow + rowCount - 1;
        int firstCol = 0;
        int lastCol = 24; // Columns A-Y (0-24)

        Workbook workbook = sheet.getWorkbook();
        CellStyle thickBorderStyle = workbook.createCellStyle();
        thickBorderStyle.cloneStyleFrom(styleService.createDataStyle(workbook));
        thickBorderStyle.setBorderTop(BorderStyle.THICK);
        thickBorderStyle.setBorderBottom(BorderStyle.THICK);
        thickBorderStyle.setBorderLeft(BorderStyle.THICK);
        thickBorderStyle.setBorderRight(BorderStyle.THICK);

        for (int r = startRow; r <= endRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = firstCol; c <= lastCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);

                CellStyle style = cell.getCellStyle();
                CellStyle newStyle = workbook.createCellStyle();
                newStyle.cloneStyleFrom(style);

                if (r == startRow) newStyle.setBorderTop(BorderStyle.THICK);
                if (r == endRow) newStyle.setBorderBottom(BorderStyle.THICK);
                if (c == firstCol) newStyle.setBorderLeft(BorderStyle.THICK);
                if (c == lastCol) newStyle.setBorderRight(BorderStyle.THICK);

                cell.setCellStyle(newStyle);
            }
        }
    }
    

    /**
     * Merges cells for IC STTN and summary data columns - UPDATED: Bold IC STTN
     */
    private void mergeCellsForIcSttn(Sheet sheet, int startRow, int rowCount, String icSttn, StationData data) {
        Row firstRow = sheet.getRow(startRow);
        Workbook workbook = sheet.getWorkbook();
        CellStyle mergedCellStyle = styleService.createMergedCellStyle(workbook);
        CellStyle boldMergedCellStyle = styleService.createBoldMergedCellStyle(workbook); // NEW: Bold style for IC STTN
        
        try {
            // ALWAYS write IC STTN name and summary data (whether merging or not)
            
            // Column A - No. of Trains (HANDEDOVER)
            Cell trainCountHandedOverCell = firstRow.getCell(0);
            if (trainCountHandedOverCell != null) {
                trainCountHandedOverCell.setCellValue(String.valueOf(data.handedOverTrainCount));
                trainCountHandedOverCell.setCellStyle(mergedCellStyle);
            }
            
            // Column B - JUMBO L+E (HANDEDOVER)
            Cell jumboHandedOverCell = firstRow.getCell(1);
            if (jumboHandedOverCell != null) {
                jumboHandedOverCell.setCellValue(formatLePlusE(data.handedOverJumboL, data.handedOverJumboE));
                jumboHandedOverCell.setCellStyle(mergedCellStyle);
            }
            
            // Column C - BOXN L+E (HANDEDOVER)
            Cell boxnHandedOverCell = firstRow.getCell(2);
            if (boxnHandedOverCell != null) {
                boxnHandedOverCell.setCellValue(formatLePlusE(data.handedOverBoxnL, data.handedOverBoxnE));
                boxnHandedOverCell.setCellStyle(mergedCellStyle);
            }
            
            // Column D - BTPN L+E (HANDEDOVER)
            Cell btpnHandedOverCell = firstRow.getCell(3);
            if (btpnHandedOverCell != null) {
                btpnHandedOverCell.setCellValue(formatLePlusE(data.handedOverBtpnL, data.handedOverBtpnE));
                btpnHandedOverCell.setCellStyle(mergedCellStyle);
            }
            
            // Column M - IC STTN (BOLD STYLE) - SHIFTED LEFT
            Cell icSttnCell = firstRow.getCell(12);
            if (icSttnCell != null) {
                icSttnCell.setCellValue(icSttn);
                icSttnCell.setCellStyle(boldMergedCellStyle); // ← USE BOLD STYLE HERE
            }
            
            // Column N - No. of Trains (TAKENOVER) - SHIFTED LEFT
            Cell trainCountTakenOverCell = firstRow.getCell(13);
            if (trainCountTakenOverCell != null) {
                trainCountTakenOverCell.setCellValue(String.valueOf(data.takenOverTrainCount));
                trainCountTakenOverCell.setCellStyle(mergedCellStyle);
            }
            
            // Column O - JUMBO L+E (TAKENOVER) - SHIFTED LEFT
            Cell jumboTakenOverCell = firstRow.getCell(14);
            if (jumboTakenOverCell != null) {
                jumboTakenOverCell.setCellValue(formatLePlusE(data.takenOverJumboL, data.takenOverJumboE));
                jumboTakenOverCell.setCellStyle(mergedCellStyle);
            }
            
            // Column P - BOXN L+E (TAKENOVER) - SHIFTED LEFT
            Cell boxnTakenOverCell = firstRow.getCell(15);
            if (boxnTakenOverCell != null) {
                boxnTakenOverCell.setCellValue(formatLePlusE(data.takenOverBoxnL, data.takenOverBoxnE));
                boxnTakenOverCell.setCellStyle(mergedCellStyle);
            }
            
            // Column Q - BTPN L+E (TAKENOVER) - SHIFTED LEFT
            Cell btpnTakenOverCell = firstRow.getCell(16);
            if (btpnTakenOverCell != null) {
                btpnTakenOverCell.setCellValue(formatLePlusE(data.takenOverBtpnL, data.takenOverBtpnE));
                btpnTakenOverCell.setCellStyle(mergedCellStyle);
            }
            
            // ONLY merge if there are multiple rows (rowCount > 1)
            if (rowCount > 1) {
                int endRow = startRow + rowCount - 1;
                
                // HANDEDOVER SECTION - Merge summary columns (A, B, C, D - indices 0, 1, 2, 3)
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 0, 0));   // No. of Trains
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));   // JUMBO L+E
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 2, 2));   // BOXN L+E
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 3, 3));   // BTPN L+E
                
                // IC STTN column (M - index 12) - SHIFTED LEFT (BOLD MERGED)
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 12, 12));
                
                // TAKENOVER SECTION - Merge summary columns (N, O, P, Q - indices 13, 14, 15, 16) - SHIFTED LEFT
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 13, 13)); // No. of Trains
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 14, 14)); // JUMBO L+E
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 15, 15)); // BOXN L+E
                sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 16, 16)); // BTPN L+E
            }
            
        } catch (Exception e) {
            System.err.println("Error processing cells for IC STTN: " + icSttn + " - " + e.getMessage());
        }
    }
    
    /**
     * Groups the classified data by IC STTN and calculates L+E counts plus details
     * UPDATED: Now handles 10 columns (8 extracted + 2 classified) with ZONE transformation
     */
    // Update this method in ExcelWriterService.java

/**
 * Groups the classified data by ZONE TO (in specific order) then by IC STTN 
 * and calculates L+E counts plus details
 * UPDATED: Now sorts by ZONE TO in order: CR, WCR, NWR, DFCR
 */
// Update groupDataByIcSttn method in ExcelWriterService.java - REMOVE zone sorting logic

/**
 * Groups the classified data by IC STTN and calculates L+E counts plus details
 * UPDATED: Data is already sorted by DataProcessingService, so just group by IC STTN
 */
private Map<String, StationData> groupDataByIcSttn(List<String[]> classifiedData) {
    // Use LinkedHashMap to preserve the insertion order (data is already sorted)
    Map<String, StationData> groupedData = new LinkedHashMap<>();
    
    // Skip header row (index 0) and process data rows
    for (int i = 1; i < classifiedData.size(); i++) {
        String[] row = classifiedData.get(i);
        
        // Extract data from the 10 columns
        String zoneTo = row[0];                           // ZONE TO (NEW)
        String icSttn = row[1];                           // IC STTN
        String handedOverSttn = row[2];                   // HANDED OVER STTN TO
        String handedOverLE = row[3];                     // HANDED OVER L/E
        String handedOverType = row[4];                   // HANDED OVER TYPE
        String handedOverClassification = row[5];         // HANDED OVER TYPE CLASSIFICATION (classified)
        String takenOverSttn = row[6];                    // TAKEN OVER STTN TO
        String takenOverLE = row[7];                      // TAKEN OVER L/E
        String takenOverType = row[8];                    // TAKEN OVER TYPE
        String takenOverClassification = row[9];          // TAKEN OVER TYPE CLASSIFICATION (classified)
        
        // APPLY ZONE TRANSFORMATION RULE
        
        
        // Get or create station data for this IC STTN (data is already zone-sorted)
        StationData stationData = groupedData.getOrDefault(icSttn, new StationData(icSttn));
        
        // Count HANDEDOVER station occurrences (ignore blanks)
        if (handedOverSttn != null && !handedOverSttn.trim().isEmpty()) {
            stationData.handedOverTrainCount++;
        }
        
        // Count TAKENOVER station occurrences (ignore blanks)
        if (takenOverSttn != null && !takenOverSttn.trim().isEmpty()) {
            stationData.takenOverTrainCount++;
        }
        
        // Count HANDEDOVER classifications by L/E
        countHandedOverData(stationData, handedOverClassification, handedOverLE);
        
        // Count TAKENOVER classifications by L/E
        countTakenOverData(stationData, takenOverClassification, takenOverLE);
        
        // Count HANDEDOVER details (L only) by classification and station
        countHandedOverDetails(stationData, handedOverClassification, handedOverLE, handedOverSttn);
        
        // Count TAKENOVER details (L only) by classification and station
        countTakenOverDetails(stationData, takenOverClassification, takenOverLE, takenOverSttn);
        
        // Count HANDEDOVER empties (E only) by wagon type - exclude CONT classification
        if (!"CONT".equalsIgnoreCase(handedOverClassification)) {
            countHandedOverEmpties(stationData, handedOverType, handedOverLE);
        }

        // Count TAKENOVER empties (E only) by wagon type - exclude CONT classification  
        if (!"CONT".equalsIgnoreCase(takenOverClassification)) {
            countTakenOverEmpties(stationData, takenOverType, takenOverLE);
        }
        
        // Count HANDEDOVER others (L only) - non-main classifications
        countHandedOverOthers(stationData, handedOverClassification, handedOverLE, handedOverSttn);
        
        // Count TAKENOVER others (L only) - non-main classifications
        countTakenOverOthers(stationData, takenOverClassification, takenOverLE, takenOverSttn);
        
        // Store in grouped data (preserves the sorted order)
        groupedData.put(icSttn, stationData);
    }
    
    System.out.println("Data grouped by IC STTN (pre-sorted by zones). Found " + groupedData.size() + " unique stations.");
    return groupedData;
}
    /**
     * NEW: Applies zone transformation rules
     * Rule: If ZONE TO = "NW" and IC STTN = "CNA", change IC STTN to "AII"
     */
    // private String applyZoneTransformation(String zoneTo, String icSttn) {
    //     if (zoneTo == null || icSttn == null) {
    //         return icSttn; // Return original if either is null
    //     }
        
    //     String cleanZoneTo = zoneTo.trim().toUpperCase();
    //     String cleanIcSttn = icSttn.trim().toUpperCase();
        
    //     // Apply transformation rule: NW + CNA → AII
    //     if ("NW".equals(cleanZoneTo) && "CNA".equals(cleanIcSttn)) {
    //         System.out.println("Zone transformation applied: ZONE TO=" + zoneTo + ", IC STTN changed from " + icSttn + " to AII");
    //         return "AII";
    //     }
        
    //     // Add more transformation rules here if needed in the future
    //     // Example:
    //     // if ("WC".equals(cleanZoneTo) && "XYZ".equals(cleanIcSttn)) {
    //     //     return "ABC";
    //     // }
        
    //     return icSttn; // Return original IC STTN if no transformation rule applies
    // }
    
    /**
     * Calculates the maximum number of stations across all classifications
     */
    private int calculateMaxStations(StationData data) {
        int maxStations = 0;
        
        // Check HANDEDOVER stations
        maxStations = Math.max(maxStations, data.handedOverJumboStations.size());
        maxStations = Math.max(maxStations, data.handedOverBoxnStations.size());
        maxStations = Math.max(maxStations, data.handedOverBtpnStations.size());
        maxStations = Math.max(maxStations, data.handedOverBtpgStations.size());
        maxStations = Math.max(maxStations, data.handedOverContStations.size());
        maxStations = Math.max(maxStations, data.handedOverShraStations.size());
        maxStations = Math.max(maxStations, data.handedOverOthersStations.size());   // NEW: Others
        maxStations = Math.max(maxStations, data.handedOverEmptiesStations.size());
        
        // Check TAKENOVER stations
        maxStations = Math.max(maxStations, data.takenOverJumboStations.size());
        maxStations = Math.max(maxStations, data.takenOverBoxnStations.size());
        maxStations = Math.max(maxStations, data.takenOverBtpnStations.size());
        maxStations = Math.max(maxStations, data.takenOverBtpgStations.size());
        maxStations = Math.max(maxStations, data.takenOverContStations.size());
        maxStations = Math.max(maxStations, data.takenOverShraStations.size());
        maxStations = Math.max(maxStations, data.takenOverOthersStations.size());    // NEW: Others
        maxStations = Math.max(maxStations, data.takenOverEmptiesStations.size());
        
        return maxStations;
    }
    
    /**
     * Writes summary data (L+E totals) to a specific row
     */
    private void writeSummaryRowData(Row excelRow, String icSttn, StationData data, CellStyle dataStyle) {
        // HANDEDOVER SECTION (Columns A-D, 0-3) - UNCHANGED
        writeCell(excelRow, 0, String.valueOf(data.handedOverTrainCount), dataStyle);
        writeCell(excelRow, 1, formatLePlusE(data.handedOverJumboL, data.handedOverJumboE), dataStyle);
        writeCell(excelRow, 2, formatLePlusE(data.handedOverBoxnL, data.handedOverBoxnE), dataStyle);
        writeCell(excelRow, 3, formatLePlusE(data.handedOverBtpnL, data.handedOverBtpnE), dataStyle);
        
        // IC STTN (Column M, index 12) - SHIFTED LEFT
        writeCell(excelRow, 12, icSttn, dataStyle);
        
        // TAKENOVER SECTION (Columns N-Q, 13-16) - SHIFTED LEFT
        writeCell(excelRow, 13, String.valueOf(data.takenOverTrainCount), dataStyle);
        writeCell(excelRow, 14, formatLePlusE(data.takenOverJumboL, data.takenOverJumboE), dataStyle);
        writeCell(excelRow, 15, formatLePlusE(data.takenOverBoxnL, data.takenOverBoxnE), dataStyle);
        writeCell(excelRow, 16, formatLePlusE(data.takenOverBtpnL, data.takenOverBtpnE), dataStyle);
    }

    /**
     * Formats L and E counts as "L+E" (e.g., "5+2"), or just "L" if E is zero, or just "E" if L is zero.
     */
    private String formatLePlusE(int l, int e) {
        if (l >= 0 && e >= 0) {
            return l + "+" + e;
        }else {
            return "";
        }
    }
    
    /**
     * Counts HANDEDOVER data by classification and L/E
     */
    private void countHandedOverData(StationData stationData, String classification, String le) {
        if (classification == null || le == null) return;
        
        // Count JUMBO by L and E
        if ("JUMBO".equalsIgnoreCase(classification.trim())) {
            if ("L".equalsIgnoreCase(le.trim())) {
                stationData.handedOverJumboL++;
            } else if ("E".equalsIgnoreCase(le.trim())) {
                stationData.handedOverJumboE++;
            }
        }
        
        // Count BOXN by L and E
        if ("BOX".equalsIgnoreCase(classification.trim()) || "BOXN".equalsIgnoreCase(classification.trim())) {
            if ("L".equalsIgnoreCase(le.trim())) {
                stationData.handedOverBoxnL++;
            } else if ("E".equalsIgnoreCase(le.trim())) {
                stationData.handedOverBoxnE++;
            }
        }
        
        // Count BTPN by L and E
        if ("BTPN".equalsIgnoreCase(classification.trim())) {
            if ("L".equalsIgnoreCase(le.trim())) {
                stationData.handedOverBtpnL++;
            } else if ("E".equalsIgnoreCase(le.trim())) {
                stationData.handedOverBtpnE++;
            }
        }
    }
    
    /**
     * Counts TAKENOVER data by classification and L/E
     */
    private void countTakenOverData(StationData stationData, String classification, String le) {
        if (classification == null || le == null) return;
        
        // Count JUMBO by L and E
        if ("JUMBO".equalsIgnoreCase(classification.trim())) {
            if ("L".equalsIgnoreCase(le.trim())) {
                stationData.takenOverJumboL++;
            } else if ("E".equalsIgnoreCase(le.trim())) {
                stationData.takenOverJumboE++;
            }
        }
        
        // Count BOXN by L and E
        if ("BOX".equalsIgnoreCase(classification.trim()) || "BOXN".equalsIgnoreCase(classification.trim())) {
            if ("L".equalsIgnoreCase(le.trim())) {
                stationData.takenOverBoxnL++;
            } else if ("E".equalsIgnoreCase(le.trim())) {
                stationData.takenOverBoxnE++;
            }
        }
        
        // Count BTPN by L and E
        if ("BTPN".equalsIgnoreCase(classification.trim())) {
            if ("L".equalsIgnoreCase(le.trim())) {
                stationData.takenOverBtpnL++;
            } else if ("E".equalsIgnoreCase(le.trim())) {
                stationData.takenOverBtpnE++;
            }
        }
    }
    
// Update this part in countHandedOverDetails method

/**
 * Counts HANDEDOVER details - station names by classification (L only, EXCEPT CONT which includes L+E)
 */
private void countHandedOverDetails(StationData stationData, String classification, String le, String station) {
    // For CONT: Include both L and E
    // For all others: Only count if L/E = "L"
    boolean shouldCount = false;
    
    if ("CONT".equalsIgnoreCase(classification != null ? classification.trim() : "")) {
        // CONT includes both L and E
        shouldCount = ("L".equalsIgnoreCase(le != null ? le.trim() : "") || 
                      "E".equalsIgnoreCase(le != null ? le.trim() : ""));
    } else {
        // All other classifications: only L
        shouldCount = "L".equalsIgnoreCase(le != null ? le.trim() : "");
    }
    
    if (!shouldCount || station == null || station.trim().isEmpty()) {
        return;
    }
    
    String cleanStation = station.trim();
    String cleanClassification = classification != null ? classification.trim() : "";
    
    // Count by classification (unchanged logic)
    if ("JUMBO".equalsIgnoreCase(cleanClassification)) {
        stationData.handedOverJumboStations.put(cleanStation, 
            stationData.handedOverJumboStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("BOX".equalsIgnoreCase(cleanClassification) || "BOXN".equalsIgnoreCase(cleanClassification)) {
        stationData.handedOverBoxnStations.put(cleanStation, 
            stationData.handedOverBoxnStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("BTPN".equalsIgnoreCase(cleanClassification)) {
        stationData.handedOverBtpnStations.put(cleanStation, 
            stationData.handedOverBtpnStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("BTPG".equalsIgnoreCase(cleanClassification)) {
        stationData.handedOverBtpgStations.put(cleanStation, 
            stationData.handedOverBtpgStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("CONT".equalsIgnoreCase(cleanClassification)) {
        // CONT now includes both L and E
        stationData.handedOverContStations.put(cleanStation, 
            stationData.handedOverContStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("SHRA".equalsIgnoreCase(cleanClassification)) {
        stationData.handedOverShraStations.put(cleanStation, 
            stationData.handedOverShraStations.getOrDefault(cleanStation, 0) + 1);
    } 
}
    
    /**
     * NEW: Counts HANDEDOVER empties - wagon types where L/E = "E" (excluding specific types)
     */
    private void countHandedOverEmpties(StationData stationData, String wagonType, String le) {
        // Only count if L/E = "E" and wagon type is not blank
        if (!"E".equalsIgnoreCase(le != null ? le.trim() : "") || 
            wagonType == null || wagonType.trim().isEmpty()) {
            return;
        }
        
        String cleanWagonType = wagonType.trim().toUpperCase();
        
        // Skip these specific wagon types - do NOT include them in empties
        if (isExcludedWagonType(cleanWagonType)) {
            return; // Ignore these wagon types
        }
        
        // Count wagon types for empties (only allowed types)
        stationData.handedOverEmptiesStations.put(cleanWagonType, 
            stationData.handedOverEmptiesStations.getOrDefault(cleanWagonType, 0) + 1);
    }
    // Update the countTakenOverEmpties method to exclude CONT from empties

/**
 * NEW: Counts TAKENOVER empties - wagon types where L/E = "E" (excluding specific types AND CONT)
 */
private void countTakenOverEmpties(StationData stationData, String wagonType, String le) {
    // Only count if L/E = "E" and wagon type is not blank
    if (!"E".equalsIgnoreCase(le != null ? le.trim() : "") || 
        wagonType == null || wagonType.trim().isEmpty()) {
        return;
    }
    
    String cleanWagonType = wagonType.trim().toUpperCase();
    
    // Skip these specific wagon types - do NOT include them in empties
    if (isExcludedWagonType(cleanWagonType)) {
        return; // Ignore these wagon types
    }
    
    // Count wagon types for empties (only allowed types)
    stationData.takenOverEmptiesStations.put(cleanWagonType, 
        stationData.takenOverEmptiesStations.getOrDefault(cleanWagonType, 0) + 1);
}
    
    // Update this part in countTakenOverDetails method

/**
 * Counts TAKENOVER details - station names by classification (L only, EXCEPT CONT which includes L+E)
 */
private void countTakenOverDetails(StationData stationData, String classification, String le, String station) {
    // For CONT: Include both L and E
    // For all others: Only count if L/E = "L"
    boolean shouldCount = false;
    
    if ("CONT".equalsIgnoreCase(classification != null ? classification.trim() : "")) {
        // CONT includes both L and E
        shouldCount = ("L".equalsIgnoreCase(le != null ? le.trim() : "") || 
                      "E".equalsIgnoreCase(le != null ? le.trim() : ""));
    } else {
        // All other classifications: only L
        shouldCount = "L".equalsIgnoreCase(le != null ? le.trim() : "");
    }

    if (!shouldCount || station == null || station.trim().isEmpty()) {
        return;
    }

    String cleanStation = station.trim();
    String cleanClassification = classification != null ? classification.trim() : "";

    // Count by classification (unchanged logic)
    if ("JUMBO".equalsIgnoreCase(cleanClassification)) {
        stationData.takenOverJumboStations.put(cleanStation,
            stationData.takenOverJumboStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("BOX".equalsIgnoreCase(cleanClassification) || "BOXN".equalsIgnoreCase(cleanClassification)) {
        stationData.takenOverBoxnStations.put(cleanStation,
            stationData.takenOverBoxnStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("BTPN".equalsIgnoreCase(cleanClassification)) {
        stationData.takenOverBtpnStations.put(cleanStation,
            stationData.takenOverBtpnStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("BTPG".equalsIgnoreCase(cleanClassification)) {
        stationData.takenOverBtpgStations.put(cleanStation,
            stationData.takenOverBtpgStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("CONT".equalsIgnoreCase(cleanClassification)) {
        // CONT now includes both L and E
        stationData.takenOverContStations.put(cleanStation,
            stationData.takenOverContStations.getOrDefault(cleanStation, 0) + 1);
    } else if ("SHRA".equalsIgnoreCase(cleanClassification)) {
        stationData.takenOverShraStations.put(cleanStation,
            stationData.takenOverShraStations.getOrDefault(cleanStation, 0) + 1);
    } 
}
    /**
     * Checks if a wagon type should be excluded from empties section
     */
    private boolean isExcludedWagonType(String wagonType) {
        // List of wagon types to exclude from empties section
        String[] excludedTypes = {
            "BFK",
            "BFKN", 
            "BKI",
            "BLC",
            "BLL",
            "BLLM",
            "BLSS",
            "BOXK"
        };
        
        for (String excludedType : excludedTypes) {
            if (excludedType.equalsIgnoreCase(wagonType)) {
                return true; // This wagon type should be excluded
            }
        }
        
        return false; // This wagon type is allowed in empties
    }
    
    /**
     * NEW: Counts HANDEDOVER others - classifications not in main categories where L/E = "L"
     */
    private void countHandedOverOthers(StationData stationData, String classification, String le, String station) {
        // Only count if L/E = "L" and both classification and station are not blank
        if (!"L".equalsIgnoreCase(le != null ? le.trim() : "") || 
            classification == null || classification.trim().isEmpty() ||
            station == null || station.trim().isEmpty()) {
            return;
        }
        
        String cleanClassification = classification.trim();
        String cleanStation = station.trim();
        
        // Check if this classification is NOT in the main 8 categories
        if (!isMainClassification(cleanClassification)) {
            // This is an "other" classification - format as "CLASSIFICATION[STATION]"
            String otherKey = cleanClassification + "[" + cleanStation + "]";
            
            stationData.handedOverOthersStations.put(otherKey, 
                stationData.handedOverOthersStations.getOrDefault(otherKey, 0) + 1);
        }
    }
    
    /**
     * NEW: Counts TAKENOVER others - classifications not in main categories where L/E = "L"
     */
    private void countTakenOverOthers(StationData stationData, String classification, String le, String station) {
        // Only count if L/E = "L" and both classification and station are not blank
        if (!"L".equalsIgnoreCase(le != null ? le.trim() : "") || 
            classification == null || classification.trim().isEmpty() ||
            station == null || station.trim().isEmpty()) {
            return;
        }
        
        String cleanClassification = classification.trim();
        String cleanStation = station.trim();
        
        // Check if this classification is NOT in the main 8 categories
        if (!isMainClassification(cleanClassification)) {
            // This is an "other" classification - format as "CLASSIFICATION[STATION]"
            String otherKey = cleanClassification + "[" + cleanStation + "]";
            
            stationData.takenOverOthersStations.put(otherKey, 
                stationData.takenOverOthersStations.getOrDefault(otherKey, 0) + 1);
        }
    }
    
    /**
     * Checks if a classification is one of the main 6 categories
     * UPDATED: Now only 6 main categories - BCACBM and NMG moved to "Others"
     */
    private boolean isMainClassification(String classification) {
        String[] mainClassifications = {
            "JUMBO",
            "BOX", "BOXN",  // BOX and BOXN are both main classifications
            "BTPN",
            "BTPG", 
            "CONT",
            "SHRA"
            // REMOVED: "BCACBM" and "NMG" - these will now be treated as "Others"
        };
        
        for (String mainClass : mainClassifications) {
            if (mainClass.equalsIgnoreCase(classification)) {
                return true; // This is a main classification
            }
        }
        
        return false; // This is NOT a main classification - it's an "other"
    }
    
    /**
     * Helper method to write a cell with proper styling
     */
    private void writeCell(Row row, int columnIndex, String value, CellStyle style) {
        Cell cell = row.createCell(columnIndex);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }
    
    /**
     * Converts workbook to byte array for download
     */
    private byte[] convertWorkbookToByteArray(Workbook workbook) throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        byte[] data = outputStream.toByteArray();
        outputStream.close();
        return data;
    }
    
    /**
     * Writes classification details (station names and counts) into the given column, starting from startRow.
     */
    private void fillClassificationColumn(Sheet sheet, int startRow, Map<String, Integer> stationMap, int columnIndex, CellStyle dataStyle) {
        int rowOffset = 0;
        for (Map.Entry<String, Integer> entry : stationMap.entrySet()) {
            Row row = sheet.getRow(startRow + rowOffset);
            if (row == null) {
                row = sheet.createRow(startRow + rowOffset);
            }
            String value = entry.getKey() + (entry.getValue() > 1 ? " (" + entry.getValue() + ")" : "");
            writeCell(row, columnIndex, value, dataStyle);
            rowOffset++;
        }
    }

    /**
     * Data class to hold counts for each IC STTN
     */
    private static class StationData {
        String icSttn;

        // Constructor to initialize icSttn
        StationData(String icSttn) {
            this.icSttn = icSttn;
        }
        
        // Train counts (station occurrences)
        int handedOverTrainCount = 0;
        int takenOverTrainCount = 0;
        
        // HANDEDOVER counts
        int handedOverJumboL = 0, handedOverJumboE = 0;
        int handedOverBoxnL = 0, handedOverBoxnE = 0;
        int handedOverBtpnL = 0, handedOverBtpnE = 0;
        
        // TAKENOVER counts
        int takenOverJumboL = 0, takenOverJumboE = 0;
        int takenOverBoxnL = 0, takenOverBoxnE = 0;
        int takenOverBtpnL = 0, takenOverBtpnE = 0;
        
        // HANDEDOVER station details (L only)
        Map<String, Integer> handedOverJumboStations = new LinkedHashMap<>();
        Map<String, Integer> handedOverBoxnStations = new LinkedHashMap<>();
        Map<String, Integer> handedOverBtpnStations = new LinkedHashMap<>();
        Map<String, Integer> handedOverBtpgStations = new LinkedHashMap<>();
        Map<String, Integer> handedOverContStations = new LinkedHashMap<>();
        Map<String, Integer> handedOverShraStations = new LinkedHashMap<>();
        Map<String, Integer> handedOverOthersStations = new LinkedHashMap<>();  // NEW: Others (L only)
        Map<String, Integer> handedOverEmptiesStations = new LinkedHashMap<>(); // Empties (E only)
        
        // TAKENOVER station details (L only)
        Map<String, Integer> takenOverJumboStations = new LinkedHashMap<>();
        Map<String, Integer> takenOverBoxnStations = new LinkedHashMap<>();
        Map<String, Integer> takenOverBtpnStations = new LinkedHashMap<>();
        Map<String, Integer> takenOverBtpgStations = new LinkedHashMap<>();
        Map<String, Integer> takenOverContStations = new LinkedHashMap<>();
        Map<String, Integer> takenOverShraStations = new LinkedHashMap<>();
        Map<String, Integer> takenOverOthersStations = new LinkedHashMap<>();   // NEW: Others (L only)
        Map<String, Integer> takenOverEmptiesStations = new LinkedHashMap<>();
      }

    // Add these methods to ExcelWriterService.java

    /**
     * Creates a bold cell style for the TOTAL row.
     */
    private CellStyle createTotalRowStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    /**
     * Adds a TOTAL row at the end of all IC STTN data
     * Calculates sum of "No of Trains" and L+E for JUMBO, BOXN, BTPN
     */
    private void addTotalRow(Sheet sheet, int nextRow, Map<String, StationData> groupedData, CellStyle dataStyle) {
        // Calculate totals
        TotalData totals = calculateTotals(groupedData);

        // Create total row at the specified index
        int totalRowIndex = nextRow;
        Row totalRow = sheet.createRow(totalRowIndex);

        // Create or use a bold style for the total row
        CellStyle totalStyle = createTotalRowStyle(sheet.getWorkbook());

        // HANDEDOVER SECTION TOTALS
        writeCell(totalRow, 0, String.valueOf(totals.handedOverTrainTotal), totalStyle);  // No. of Trains
        writeCell(totalRow, 1, formatLePlusE(totals.handedOverJumboL, totals.handedOverJumboE), totalStyle);  // JUMBO L+E
        writeCell(totalRow, 2, formatLePlusE(totals.handedOverBoxnL, totals.handedOverBoxnE), totalStyle);   // BOXN L+E
        writeCell(totalRow, 3, formatLePlusE(totals.handedOverBtpnL, totals.handedOverBtpnE), totalStyle);   // BTPN L+E

        // IC STTN column - Write "TOTAL" in BOLD
        writeCell(totalRow, 12, "TOTAL", totalStyle);

        // TAKENOVER SECTION TOTALS
        writeCell(totalRow, 13, String.valueOf(totals.takenOverTrainTotal), totalStyle);   // No. of Trains
        writeCell(totalRow, 14, formatLePlusE(totals.takenOverJumboL, totals.takenOverJumboE), totalStyle);   // JUMBO L+E
        writeCell(totalRow, 15, formatLePlusE(totals.takenOverBoxnL, totals.takenOverBoxnE), totalStyle);    // BOXN L+E
        writeCell(totalRow, 16, formatLePlusE(totals.takenOverBtpnL, totals.takenOverBtpnE), totalStyle);    // BTPN L+E

        System.out.println("TOTAL row added at row " + (totalRowIndex + 1) + " with calculated sums.");
    }

    /**
     * Calculates totals across all IC STTN stations
     */
    private TotalData calculateTotals(Map<String, StationData> groupedData) {
        TotalData totals = new TotalData();
        
        for (StationData stationData : groupedData.values()) {
            // HANDEDOVER totals
            totals.handedOverTrainTotal += stationData.handedOverTrainCount;
            totals.handedOverJumboL += stationData.handedOverJumboL;
            totals.handedOverJumboE += stationData.handedOverJumboE;
            totals.handedOverBoxnL += stationData.handedOverBoxnL;
            totals.handedOverBoxnE += stationData.handedOverBoxnE;
            totals.handedOverBtpnL += stationData.handedOverBtpnL;
            totals.handedOverBtpnE += stationData.handedOverBtpnE;
            
            // TAKENOVER totals
            totals.takenOverTrainTotal += stationData.takenOverTrainCount;
            totals.takenOverJumboL += stationData.takenOverJumboL;
            totals.takenOverJumboE += stationData.takenOverJumboE;
            totals.takenOverBoxnL += stationData.takenOverBoxnL;
            totals.takenOverBoxnE += stationData.takenOverBoxnE;
            totals.takenOverBtpnL += stationData.takenOverBtpnL;
            totals.takenOverBtpnE += stationData.takenOverBtpnE;
        }
        
        System.out.println("Calculated totals: HANDEDOVER Trains=" + totals.handedOverTrainTotal + 
                          ", TAKENOVER Trains=" + totals.takenOverTrainTotal);
        return totals;
    }

    /**
     * Helper class to store total calculations
     */
    private static class TotalData {
        int handedOverTrainTotal = 0;
        int handedOverJumboL = 0, handedOverJumboE = 0;
        int handedOverBoxnL = 0, handedOverBoxnE = 0;
        int handedOverBtpnL = 0, handedOverBtpnE = 0;
        
        int takenOverTrainTotal = 0;
        int takenOverJumboL = 0, takenOverJumboE = 0;
        int takenOverBoxnL = 0, takenOverBoxnE = 0;
        int takenOverBtpnL = 0, takenOverBtpnE = 0;
    }

    /**
     * Main method to write all data to Excel sheet
     */
    public void writeToExcel(Sheet sheet, List<String[]> classifiedData) {
        // Group data by IC STTN and calculate L+E counts
        Map<String, StationData> groupedData = groupDataByIcSttn(classifiedData);
        
        // Create styles
        Workbook workbook = sheet.getWorkbook();
        CellStyle dataStyle = styleService.createDataStyle(workbook);
        
        // Write data starting from row 5 (index 4, after header rows 1-4)
        int currentRow = 4;
        
        for (Map.Entry<String, StationData> entry : groupedData.entrySet()) {
            String icSttn = entry.getKey();
            StationData stationData = entry.getValue();
            
            // Write complete data for this IC STTN (returns next available row)
            currentRow = writeCompleteStationData(sheet, currentRow, icSttn, stationData, dataStyle);
        }
        
        // NEW: Add TOTAL row after all data
        addTotalRow(sheet, currentRow, groupedData, dataStyle);
        
        System.out.println("Excel writing completed. Data written for " + groupedData.size() + " IC STTN stations.");
    }

    // Add this method to ExcelWriterService.java

    /**
     * Creates a simple 10-column Excel file showing the intermediate processed data
     * This is for debugging/verification purposes
     */
    public byte[] createIntermediateExcel(List<String[]> classifiedData) throws Exception {
        // Create Excel workbook
        Workbook workbook = new XSSFWorkbook();
        
        // Create the intermediate data sheet
        Sheet intermediateSheet = workbook.createSheet("Intermediate 10-Column Data");
        
        // Create basic styles
        CellStyle headerStyle = createIntermediateHeaderStyle(workbook);
        CellStyle dataStyle = createIntermediateDataStyle(workbook);
        
        // Write header row
        if (!classifiedData.isEmpty()) {
            Row headerRow = intermediateSheet.createRow(0);
            String[] headers = classifiedData.get(0);
            
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }
        }
        
        // Write data rows
        for (int i = 1; i < classifiedData.size(); i++) {
            Row dataRow = intermediateSheet.createRow(i);
            String[] rowData = classifiedData.get(i);
            
            for (int j = 0; j < rowData.length; j++) {
                Cell cell = dataRow.createCell(j);
                cell.setCellValue(rowData[j] != null ? rowData[j] : "");
                cell.setCellStyle(dataStyle);
            }
        }
        
        // Auto-size columns for better readability
        for (int i = 0; i < 10; i++) {
            intermediateSheet.autoSizeColumn(i);
            // Set minimum width to prevent too narrow columns
            if (intermediateSheet.getColumnWidth(i) < 3000) {
                intermediateSheet.setColumnWidth(i, 3000);
            }
        }
        
        // Convert to byte array
        byte[] excelData = convertWorkbookToByteArray(workbook);
        
        // Close workbook
        workbook.close();
        
        System.out.println("✅ Intermediate 10-column Excel created with " + (classifiedData.size() - 1) + " data rows.");
        return excelData;
    }

    /**
     * Creates header style for intermediate Excel
     */
    private CellStyle createIntermediateHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        
        // Font
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        font.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(font);
        
        // Background
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Alignment
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        return style;
    }

    /**
     * Creates data style for intermediate Excel
     */
    private CellStyle createIntermediateDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        
        // Font
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        
        // Alignment
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        return style;
    }
}