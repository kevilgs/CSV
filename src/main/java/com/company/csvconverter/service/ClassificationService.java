package com.company.csvconverter.service;

import com.company.csvconverter.controller.ClassificationController.ClassificationRequest;
import com.opencsv.CSVReader;
import com.opencsv.CSVWriter;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

@Service
public class ClassificationService {
    
    private static final String CLASSIFICATION_FILE = "wagon_classifications.csv";
    private static final String CLASSIFICATION_DIR = "data";
    
    // Cache for classifications
    private Map<String, String> classificationCache = new HashMap<>();
    private boolean cacheLoaded = false;
    
    public int saveNewClassifications(List<ClassificationRequest> newClassifications) throws Exception {
        // Load existing classifications
        Map<String, String> existingClassifications = loadClassificationsFromFile();
        
        int savedCount = 0;
        List<String[]> allClassifications = new ArrayList<>();
        
        // Add existing classifications to list
        for (Map.Entry<String, String> entry : existingClassifications.entrySet()) {
            allClassifications.add(new String[]{entry.getKey(), entry.getValue()});
        }
        
        // Process new classifications
        for (ClassificationRequest request : newClassifications) {
            String category = request.getCategory().toUpperCase();
            
            for (String wagonType : request.getWagonTypes()) {
                String cleanWagonType = wagonType.trim().toUpperCase();
                
                // Check if this wagon type already exists
                if (!existingClassifications.containsKey(cleanWagonType)) {
                    allClassifications.add(new String[]{cleanWagonType, category});
                    savedCount++;
                    System.out.println("Added new classification: " + cleanWagonType + " -> " + category);
                } else {
                    System.out.println("Classification already exists: " + cleanWagonType + " -> " + existingClassifications.get(cleanWagonType));
                }
            }
        }
        
        // Save all classifications back to file
        saveClassificationsToFile(allClassifications);
        
        // Clear cache to force reload
        cacheLoaded = false;
        classificationCache.clear();
        
        return savedCount;
    }
    
    public Map<String, String> getAllClassifications() throws Exception {
        if (!cacheLoaded) {
            classificationCache = loadClassificationsFromFile();
            cacheLoaded = true;
        }
        return new HashMap<>(classificationCache);
    }
    
    public String getClassification(String wagonType) {
        try {
            if (!cacheLoaded) {
                classificationCache = loadClassificationsFromFile();
                cacheLoaded = true;
            }
            return classificationCache.get(wagonType.trim().toUpperCase());
        } catch (Exception e) {
            System.err.println("Error getting classification for " + wagonType + ": " + e.getMessage());
            return null;
        }
    }
    
    private Map<String, String> loadClassificationsFromFile() throws Exception {
        Map<String, String> classifications = new HashMap<>();
        
        // First load default classifications
        loadDefaultClassifications(classifications);
        
        // Then load from file if it exists
        Path filePath = getClassificationFilePath();
        if (Files.exists(filePath)) {
            try (CSVReader reader = new CSVReader(new FileReader(filePath.toFile()))) {
                String[] line;
                reader.readNext(); // Skip header if present
                
                while ((line = reader.readNext()) != null) {
                    if (line.length >= 2 && !line[0].trim().isEmpty()) {
                        classifications.put(line[0].trim().toUpperCase(), line[1].trim().toUpperCase());
                    }
                }
            }
        } else {
            // Create file with default classifications
            saveClassificationsToFile(convertMapToList(classifications));
        }
        
        System.out.println("Loaded " + classifications.size() + " wagon classifications");
        return classifications;
    }
    
    private void loadDefaultClassifications(Map<String, String> classifications) {
        // JUMBO category
        String[] jumbo = {"BCN", "BCNAHSM1", "BCNAHSM2", "BCNHL", "BCNM"};
        for (String type : jumbo) {
            classifications.put(type, "JUMBO");
        }
        
        // BOX category
        String[] box = {"BOXN", "BOXNEL", "BOXNHL", "BOXNHL25T", "BOXNR", "BOXNS", "BOXNER"};
        for (String type : box) {
            classifications.put(type, "BOX");
        }
        
        // BTPN category
        String[] btpn = {"BTPN", "BTFNL"};
        for (String type : btpn) {
            classifications.put(type, "BTPN");
        }
        
        // BTPG category
        String[] btpg = {"BTPG", "BTPGN"};
        for (String type : btpg) {
            classifications.put(type, "BTPG");
        }
        
        // CONT category
        String[] cont = {"BFK", "BFKN", "BKI", "BLC", "BLL", "BLLM", "BLSS", "BOXK"};
        for (String type : cont) {
            classifications.put(type, "CONT");
        }
        
        // SHRA category
        String[] shra = {"BFNS", "BFNS22.9", "BFNSM", "BFNSM1", "BFNV", "BRN", "BRN22.9", "SHRA", "SHRN", "BOST", "BOSM"};
        for (String type : shra) {
            classifications.put(type, "SHRA");
        }
        
        // Single item categories
        classifications.put("BCACBM", "BCACBM");
        classifications.put("NMG", "NMG");
        classifications.put("NMGHS", "NMG");
        classifications.put("ACT1", "ACT1");
        classifications.put("BCBFG", "BCBFG");
        classifications.put("BCFCM", "BCFCM");
        classifications.put("MYLY", "MYLY");
    }
    
    private void saveClassificationsToFile(List<String[]> classifications) throws Exception {
        Path filePath = getClassificationFilePath();
        
        // Create directory if it doesn't exist
        Files.createDirectories(filePath.getParent());
        
        try (CSVWriter writer = new CSVWriter(new FileWriter(filePath.toFile()))) {
            // Write header
            writer.writeNext(new String[]{"WAGON_TYPE", "CATEGORY"});
            
            // Sort classifications for better readability
            classifications.sort(Comparator.comparing(arr -> arr[0]));
            
            // Write all classifications
            writer.writeAll(classifications);
        }
        
        System.out.println("Saved " + classifications.size() + " classifications to " + filePath);
    }
    
    private List<String[]> convertMapToList(Map<String, String> map) {
        List<String[]> list = new ArrayList<>();
        for (Map.Entry<String, String> entry : map.entrySet()) {
            list.add(new String[]{entry.getKey(), entry.getValue()});
        }
        return list;
    }
    
    private Path getClassificationFilePath() {
        return Paths.get(CLASSIFICATION_DIR, CLASSIFICATION_FILE);
    }
}
