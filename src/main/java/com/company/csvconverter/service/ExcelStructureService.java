package com.company.csvconverter.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class ExcelStructureService {
    
    @Autowired
    private ExcelStyleService styleService;
    
    /**
     * Creates the complete report structure (Rows 1-4)
     * Row 1: Main Title
     * Row 2: Section Headers (HANDEDOVER / TAKENOVER)
     * Row 3: Column Headers (No. of Trains, JUMBO, etc.)
     * Row 4: Sub-column Headers (L+E, detailed classifications)
     */
    public void createReportStructure(Sheet sheet, Workbook workbook) {
        // Create all styles
        CellStyle titleStyle = styleService.createTitleStyle(workbook);
        CellStyle sectionHeaderStyle = styleService.createSectionHeaderStyle(workbook);
        CellStyle columnHeaderStyle = styleService.createColumnHeaderStyle(workbook);
        CellStyle subColumnHeaderStyle = styleService.createColumnHeaderStyle(workbook); // Same style for now
        
        // Create the structure
        createMainTitle(sheet, titleStyle);
        createSectionHeaders(sheet, sectionHeaderStyle);
        createColumnHeaders(sheet, columnHeaderStyle);
        createSubColumnHeaders(sheet, subColumnHeaderStyle); // NEW ROW 4
        setColumnWidths(sheet);
    }
    
    /**
     * Creates Row 1: "Zonal Interchange as ON" spanning A1 to Y1
     * UPDATED: Reduced span to match new compact structure
     */
    private void createMainTitle(Sheet sheet, CellStyle titleStyle) {
        Row titleRow = sheet.createRow(0);
        
        // Create and style ALL cells in the merged range (A1 to Y1) - COMPACT
        for (int col = 0; col <= 24; col++) {
            Cell cell = titleRow.createCell(col);
            if (col == 0) {
                cell.setCellValue("Zonal Interchange as ON");
            }
            cell.setCellStyle(titleStyle);
        }
        
        // Merge cells from A1 to Y1 (25 columns - A=0 to Y=24) - COMPACT
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 24)); // A1:Y1
    }
    
    /**
     * Creates Row 2: Section headers - UPDATED: Compact structure with no empty columns
     */
    private void createSectionHeaders(Sheet sheet, CellStyle sectionHeaderStyle) {
        Row sectionRow = sheet.createRow(1);
        
        // HANDEDOVER section (A2 to L2)
        for (int col = 0; col <= 11; col++) {
            Cell cell = sectionRow.createCell(col);
            if (col == 0) {
                cell.setCellValue("HANDEDOVER");
            }
            cell.setCellStyle(sectionHeaderStyle);
        }
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 11)); // A2:L2
        
        // IC STTN column (M2 - index 12) - SHIFTED LEFT
        Cell icSttnCell = sectionRow.createCell(12);
        icSttnCell.setCellValue("IC STTN");
        icSttnCell.setCellStyle(sectionHeaderStyle);
        
        // TAKENOVER section (N2 to Y2 - columns 13-24) - SHIFTED LEFT
        for (int col = 13; col <= 24; col++) {
            Cell cell = sectionRow.createCell(col);
            if (col == 13) {
                cell.setCellValue("TAKENOVER");
            }
            cell.setCellStyle(sectionHeaderStyle);
        }
        sheet.addMergedRegion(new CellRangeAddress(1, 1, 13, 24)); // N2:Y2
    }
    
    /**
     * Creates Row 3: Column headers - UPDATED: Shifted IC STTN and TAKENOVER left
     */
    private void createColumnHeaders(Sheet sheet, CellStyle columnHeaderStyle) {
        Row headerRow = sheet.createRow(2);
        
        // HANDEDOVER section columns (A3 to L3)
        createHandedOverColumns(headerRow, columnHeaderStyle);
        
        // IC STTN column (M3 - index 12) - SHIFTED LEFT
        Cell icSttnHeaderCell = headerRow.createCell(12);
        icSttnHeaderCell.setCellValue("IC STTN");
        icSttnHeaderCell.setCellStyle(columnHeaderStyle);
        
        // TAKENOVER section columns (N3 to Y3) - SHIFTED LEFT
        createTakenOverColumns(headerRow, columnHeaderStyle);
    }
    
    /**
     * Creates HANDEDOVER section columns (A3 to L3) - UPDATED: Removed empty columns M-Q
     */
    private void createHandedOverColumns(Row headerRow, CellStyle columnHeaderStyle) {
        String[] handedOverColumns = {"No. of Trains", "JUMBO", "BOXN", "BTPN", "DETAILS"};
        int col = 0;
        
        // Add HANDEDOVER individual columns (A, B, C, D)
        for (int i = 0; i < 4; i++) { // No. of Trains, JUMBO, BOXN, BTPN
            Cell cell = headerRow.createCell(col++);
            cell.setCellValue(handedOverColumns[i]);
            cell.setCellStyle(columnHeaderStyle);
        }
        
        // DETAILS spans from column E to L (8 columns: JUMBO, BOXN, BTPN, BTPG, CONT, SHRA, Others, Empties)
        for (int detailCol = 4; detailCol <= 11; detailCol++) {
            Cell cell = headerRow.createCell(detailCol);
            if (detailCol == 4) {
                cell.setCellValue("DETAILS");
            }
            cell.setCellStyle(columnHeaderStyle);
        }
        headerRow.getSheet().addMergedRegion(new CellRangeAddress(2, 2, 4, 11)); // E3:L3
        
        // SKIP columns M-Q (12-16) - these will remain empty/unused
    }
    
    /**
     * Creates TAKENOVER section columns (N3 to Y3) - UPDATED: Shifted left to start at column 13
     */
    private void createTakenOverColumns(Row headerRow, CellStyle columnHeaderStyle) {
        String[] takenOverColumns = {"No. of Trains", "JUMBO", "BOXN", "BTPN", "DETAILS"};
        int col = 13; // Start from column N (index 13) - SHIFTED LEFT
        
        // Add TAKENOVER individual columns (N, O, P, Q)
        for (int i = 0; i < 4; i++) { // No. of Trains, JUMBO, BOXN, BTPN
            Cell cell = headerRow.createCell(col++);
            cell.setCellValue(takenOverColumns[i]);
            cell.setCellStyle(columnHeaderStyle);
        }
        
        // DETAILS spans from column R to Y (8 columns) - SHIFTED LEFT
        for (int detailCol = 17; detailCol <= 24; detailCol++) {
            Cell cell = headerRow.createCell(detailCol);
            if (detailCol == 17) {
                cell.setCellValue("DETAILS");
            }
            cell.setCellStyle(columnHeaderStyle);
        }
        headerRow.getSheet().addMergedRegion(new CellRangeAddress(2, 2, 17, 24)); // R3:Y3
    }
    
    /**
     * Creates Row 4: Sub-column headers with detailed classifications
     */
    private void createSubColumnHeaders(Sheet sheet, CellStyle subColumnHeaderStyle) {
        Row subHeaderRow = sheet.createRow(3); // Row 4 (index 3)
        
        // HANDEDOVER section sub-headers
        createHandedOverSubHeaders(subHeaderRow, subColumnHeaderStyle);
        
        // IC STTN remains merged from Row 3
        mergeIcSttnWithRow4(sheet, subColumnHeaderStyle);
        
        // TAKENOVER section sub-headers  
        createTakenOverSubHeaders(subHeaderRow, subColumnHeaderStyle);
    }
    
    /**
     * Creates HANDEDOVER sub-headers (Row 4, Columns A-L) - UPDATED: Only fill used columns
     */
    private void createHandedOverSubHeaders(Row subHeaderRow, CellStyle subColumnHeaderStyle) {
        // Columns B, C, D: L+E under JUMBO, BOXN, BTPN
        String[] lePlusEColumns = {"L+E", "L+E", "L+E"};
        for (int i = 0; i < 3; i++) {
            Cell cell = subHeaderRow.createCell(i + 1); // Columns B, C, D
            cell.setCellValue(lePlusEColumns[i]);
            cell.setCellStyle(subColumnHeaderStyle);
        }
        
        // Columns E-L: UPDATED - Only 6 main categories + Others + Empties (8 total)
        String[] detailsSubCategories = {
            "JUMBO",    // Column E (4)
            "BOXN",     // Column F (5)
            "BTPN",     // Column G (6)
            "BTPG",     // Column H (7)
            "CONT",     // Column I (8)
            "SHRA",     // Column J (9)
            "Others",   // Column K (10)
            "Empties"   // Column L (11)
        };
        
        int col = 4; // Start from column E
        for (String category : detailsSubCategories) {
            Cell cell = subHeaderRow.createCell(col++);
            cell.setCellValue(category);
            cell.setCellStyle(subColumnHeaderStyle);
        }
        
        // SKIP columns M-Q (12-16) - leave them empty
    }
    
    /**
     * Creates TAKENOVER sub-headers (Row 4, Columns N-Y) - UPDATED: Shifted left
     */
    private void createTakenOverSubHeaders(Row subHeaderRow, CellStyle subColumnHeaderStyle) {
        // Columns O, P, Q: L+E under JUMBO, BOXN, BTPN - SHIFTED LEFT
        String[] lePlusEColumns = {"L+E", "L+E", "L+E"};
        for (int i = 0; i < 3; i++) {
            Cell cell = subHeaderRow.createCell(i + 14); // Columns O, P, Q (indices 14, 15, 16)
            cell.setCellValue(lePlusEColumns[i]);
            cell.setCellStyle(subColumnHeaderStyle);
        }
        
        // Columns R-Y: DETAILS subcategories - SHIFTED LEFT
        String[] detailsSubCategories = {
            "JUMBO",    // Column R (17)
            "BOXN",     // Column S (18)
            "BTPN",     // Column T (19)
            "BTPG",     // Column U (20)
            "CONT",     // Column V (21)
            "SHRA",     // Column W (22)
            "Others",   // Column X (23)
            "Empties"   // Column Y (24)
        };
        
        int col = 17; // Start from column R (index 17) - SHIFTED LEFT
        for (String category : detailsSubCategories) {
            Cell cell = subHeaderRow.createCell(col++);
            cell.setCellValue(category);
            cell.setCellStyle(subColumnHeaderStyle);
        }
    }
    
    /**
     * Merges "No. of Trains" and "IC STTN" cells from Row 3 to Row 4 - UPDATED: Shifted positions
     */
    private void mergeIcSttnWithRow4(Sheet sheet, CellStyle style) {
        // Merge HANDEDOVER "No. of Trains" (Column A, Rows 3-4)
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 0)); // A3:A4
        
        // Merge IC STTN (Column M, Rows 3-4) - SHIFTED LEFT
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 12, 12)); // M3:M4
        
        // Merge TAKENOVER "No. of Trains" (Column N, Rows 3-4) - SHIFTED LEFT
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 13, 13)); // N3:N4
        
        // Apply style to Row 4 cells that are part of merged regions
        Row subHeaderRow = sheet.getRow(3);
        if (subHeaderRow != null) {
            // Create styled cells for merged regions in Row 4
            Cell cellA4 = subHeaderRow.createCell(0);
            cellA4.setCellStyle(style);
            
            Cell cellM4 = subHeaderRow.createCell(12); // SHIFTED LEFT
            cellM4.setCellStyle(style);
            
            Cell cellN4 = subHeaderRow.createCell(13); // SHIFTED LEFT
            cellN4.setCellStyle(style);
        }
    }
    
    /**
     * Sets optimized column widths - UPDATED: New column positions
     */
    private void setColumnWidths(Sheet sheet) {
        // Set width for "No. of Trains" columns
        sheet.setColumnWidth(0, 4000);  // HANDEDOVER "No. of Trains" (column A)
        sheet.setColumnWidth(13, 4000); // TAKENOVER "No. of Trains" (column N) - SHIFTED LEFT
        
        // Set width for "IC STTN" column
        sheet.setColumnWidth(12, 3500); // IC STTN (column M) - SHIFTED LEFT
        
        // Set standard width for JUMBO, BOXN, BTPN columns
        sheet.setColumnWidth(1, 2500);  // HANDEDOVER JUMBO
        sheet.setColumnWidth(2, 2500);  // HANDEDOVER BOXN
        sheet.setColumnWidth(3, 2500);  // HANDEDOVER BTPN
        
        sheet.setColumnWidth(14, 2500); // TAKENOVER JUMBO (column O) - SHIFTED LEFT
        sheet.setColumnWidth(15, 2500); // TAKENOVER BOXN (column P) - SHIFTED LEFT
        sheet.setColumnWidth(16, 2500); // TAKENOVER BTPN (column Q) - SHIFTED LEFT
        
        // Auto-size the HANDEDOVER DETAILS sections (E-L)
        for (int col = 4; col <= 11; col++) { // HANDEDOVER DETAILS (E-L)
            sheet.autoSizeColumn(col);
            if (sheet.getColumnWidth(col) < 2000) {
                sheet.setColumnWidth(col, 2000);
            }
        }
        
        // Auto-size the TAKENOVER DETAILS sections (R-Y) - SHIFTED LEFT
        for (int col = 17; col <= 24; col++) { // TAKENOVER DETAILS (R-Y)
            sheet.autoSizeColumn(col);
            if (sheet.getColumnWidth(col) < 2000) {
                sheet.setColumnWidth(col, 2000);
            }
        }
    }
}
