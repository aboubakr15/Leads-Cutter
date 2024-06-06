package com.cutter.cutter;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import java.io.File;
import java.io.IOException;

public class Config {
    public File DEFAULT_SAVE_BASE_DIR_X = new File("");
    public File DEFAULT_SAVE_BASE_DIR_N = new File("");
    public File DEFAULT_SAVE_RESULT_DIR_1 = new File("");
    public File DEFAULT_SAVE_RESULT_DIR_2 = new File("");
    public File DEFAULT_SAVE_RESULT_DIR_3 = new File("");
    public File DEFAULT_SAV_EMAILS_DIR = new File("");
    public File DEFAULT_INVENTORY_DIR = new File("");

    Config() {
        File configFile = new File("./config/config.json");

        if (configFile.exists()) {
            ObjectMapper objectMapper = new ObjectMapper();

            try {
                JsonNode rootNode = objectMapper.readTree(configFile);

                DEFAULT_SAV_EMAILS_DIR = new File(rootNode.get("DEFAULT_SAV_EMAILS_DIR").asText());
                DEFAULT_SAVE_BASE_DIR_N = new File(rootNode.get("DEFAULT_SAVE_BASE_DIR_N").asText());
                DEFAULT_SAVE_BASE_DIR_X = new File(rootNode.get("DEFAULT_SAVE_BASE_DIR_X").asText());
                DEFAULT_SAVE_RESULT_DIR_1 = new File(rootNode.get("DEFAULT_SAVE_RESULT_DIR_1").asText());
                DEFAULT_SAVE_RESULT_DIR_2 = new File(rootNode.get("DEFAULT_SAVE_RESULT_DIR_2").asText());
                DEFAULT_SAVE_RESULT_DIR_3 = new File(rootNode.get("DEFAULT_SAVE_RESULT_DIR_3").asText());
                DEFAULT_INVENTORY_DIR = new File(rootNode.get("DEFAULT_INVENTORY_DIR").asText());

            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            save();
        }
    }

    public void save() {
        try {
            ObjectMapper objectMapper = new ObjectMapper();
            ObjectNode rootNode = objectMapper.createObjectNode();

            rootNode.put("DEFAULT_SAVE_BASE_DIR_N", DEFAULT_SAVE_BASE_DIR_N.getAbsolutePath());
            rootNode.put("DEFAULT_SAVE_BASE_DIR_X", DEFAULT_SAVE_BASE_DIR_X.getAbsolutePath());
            rootNode.put("DEFAULT_SAVE_RESULT_DIR_1", DEFAULT_SAVE_RESULT_DIR_1.getAbsolutePath());
            rootNode.put("DEFAULT_SAVE_RESULT_DIR_2", DEFAULT_SAVE_RESULT_DIR_2.getAbsolutePath());
            rootNode.put("DEFAULT_SAVE_RESULT_DIR_3", DEFAULT_SAVE_RESULT_DIR_3.getAbsolutePath());
            rootNode.put("DEFAULT_SAV_EMAILS_DIR", DEFAULT_SAV_EMAILS_DIR.getAbsolutePath());
            rootNode.put("DEFAULT_INVENTORY_DIR", DEFAULT_INVENTORY_DIR.getAbsolutePath());

            File configFile = new File("./config/config.json");

            if (!configFile.getParentFile().exists()) {
                configFile.getParentFile().mkdirs();
            }

            if(!configFile.exists()) {
                configFile.createNewFile();
            }

            objectMapper.writeValue(configFile, rootNode);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
