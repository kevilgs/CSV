package com.company.csvconverter.controller;

import com.company.csvconverter.service.ClassificationService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api")
public class ClassificationController {
    
    @Autowired
    private ClassificationService classificationService;
    
    @PostMapping("/classifications")
    public ResponseEntity<Map<String, Object>> saveClassifications(@RequestBody List<ClassificationRequest> classifications) {
        Map<String, Object> response = new HashMap<>();
        
        try {
            int saved = classificationService.saveNewClassifications(classifications);
            response.put("success", true);
            response.put("message", "Successfully saved " + saved + " new classifications");
            response.put("savedCount", saved);
            
            return ResponseEntity.ok(response);
            
        } catch (Exception e) {
            response.put("success", false);
            response.put("message", "Failed to save classifications: " + e.getMessage());
            
            return ResponseEntity.badRequest().body(response);
        }
    }
    
    @GetMapping("/classifications")
    public ResponseEntity<Map<String, String>> getAllClassifications() {
        try {
            Map<String, String> classifications = classificationService.getAllClassifications();
            return ResponseEntity.ok(classifications);
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(new HashMap<>());
        }
    }
    
    // Inner class for request mapping
    public static class ClassificationRequest {
        private String category;
        private List<String> wagonTypes;
        
        // Getters and setters
        public String getCategory() { return category; }
        public void setCategory(String category) { this.category = category; }
        public List<String> getWagonTypes() { return wagonTypes; }
        public void setWagonTypes(List<String> wagonTypes) { this.wagonTypes = wagonTypes; }
    }
}
