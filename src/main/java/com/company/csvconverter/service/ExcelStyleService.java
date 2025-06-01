package com.company.csvconverter.service;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

@Service
public class ExcelStyleService {
    
    /**
     * Creates the main title style (Row 1)
     * Light grey background, black Comic Sans 13pt font
     */
    public CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setFontName("Comic Sans MS");
        style.setFont(font);
        
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Borders
        style.setBorderTop(BorderStyle.THICK);
        style.setBorderBottom(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        
        return style;
    }
    
    /**
     * Creates section header style (Row 2)
     * HANDEDOVER / TAKENOVER sections
     */
    public CellStyle createSectionHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setFontName("Comic Sans MS");
        style.setFont(font);
        
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
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
     * Creates column header style (Row 3)
     * No. of Trains, JUMBO, BOXN, etc.
     */
    public CellStyle createColumnHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setFontName("Comic Sans MS");
        style.setFont(font);
        
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
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
     * Creates data cell style (for data rows)
     */
    public CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Comic Sans MS");
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        
        // Borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        return style;
    }
    
    /**
     * Creates classification cell style (highlighted data)
     */
    public CellStyle createClassificationStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Comic Sans MS");
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        
        // Light yellow background for classification columns
        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        return style;
    }
    
    /**
     * Creates a style for merged cells with center alignment
     */
    public CellStyle createMergedCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        
        // Set alignment
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Set borders
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        // Set font
        Font font = workbook.createFont();
        font.setFontName("Comic Sans MS");
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        
        return style;
    }
    
    /**
     * Creates a style for thick outside borders around IC STTN blocks
     */
    public CellStyle createThickBorderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        
        // Set alignment
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Set THICK borders
        style.setBorderTop(BorderStyle.THICK);
        style.setBorderBottom(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        
        // Set border colors (optional - can be black)
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        
        // Set font
        Font font = workbook.createFont();
        font.setFontName("Comic Sans MS");
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        
        return style;
    }
    
    /**
     * Creates a BOLD version of the merged cell style specifically for IC STTN values
     * Clones the existing merged cell style and adds bold font
     */
    public CellStyle createBoldMergedCellStyle(Workbook workbook) {
        // Get the base merged cell style
        CellStyle baseStyle = createMergedCellStyle(workbook);
        
        // Create a new style and clone all properties from base style
        CellStyle boldStyle = workbook.createCellStyle();
        boldStyle.cloneStyleFrom(baseStyle);
        
        // Create a BOLD font (copy existing font properties and make it bold)
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Comic Sans MS");
        boldFont.setFontHeightInPoints((short) 12);
        boldFont.setBold(true);  // ← MAKE IT BOLD
        boldFont.setColor(IndexedColors.BLACK.getIndex());
        
        // Apply the bold font to the style
        boldStyle.setFont(boldFont);
        
        return boldStyle;
    }
    
    /**
     * Creates a BOLD style for the TOTAL row with borders
     */
    public CellStyle createTotalRowStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        
        // Set alignment
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        // Set BOLD font for TOTAL row
        Font boldFont = workbook.createFont();
        boldFont.setFontName("Comic Sans MS");
        boldFont.setFontHeightInPoints((short) 12);
        boldFont.setBold(true);  // ← MAKE IT BOLD
        boldFont.setColor(IndexedColors.BLACK.getIndex());
        style.setFont(boldFont);
        
        // Light background to distinguish TOTAL row
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        
        // Set THICK borders for TOTAL row
        style.setBorderTop(BorderStyle.THICK);
        style.setBorderBottom(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        
        return style;
    }
}
