package com.uitc;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ConfigLoader {
    private final List<String> keywords;
    private final List<String> datePatterns;

    private final Map<String, String> patternByReport;
    
    public ConfigLoader(String jsonConfigPath) throws IOException {
        ObjectMapper mapper = new ObjectMapper();

        try (InputStream is = ConfigLoader.class.getResourceAsStream(jsonConfigPath)) {
            if(is == null) {
                throw new IllegalArgumentException("找不到 .json 檔");
            }

            JsonNode root = mapper.readTree(is);

            this.keywords = new ArrayList<>();
            this.datePatterns = new ArrayList<>();

            root.get("keywords").forEach(node -> keywords.add(node.asText()));
            root.get("date_patterns").forEach(node -> datePatterns.add(node.asText()));

            this.patternByReport = new HashMap<>();

            JsonNode patternNode = root.get("pattern_by_report");
            if (patternNode != null) {
                patternNode.fields().forEachRemaining(entry -> {
                    patternByReport.put(entry.getKey(), entry.getValue().asText());
                });
            }
        }
    }

    public List<String> getKeywords() {
        return keywords;
    }

    public List<String> getDatePatterns() {
        return datePatterns;
    }

    public Map<String, String> getPatternByReport() {
        return patternByReport;
    }
}
