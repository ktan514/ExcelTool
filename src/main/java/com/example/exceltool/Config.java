package com.example.exceltool;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class Config {
    private final Properties prop;
    private static Config instance = null;

    public Config() throws IOException {
        this.prop = new Properties();
        String filename = "./config.ini";
        File file = new File(filename);
        InputStream input = new FileInputStream(file);
        this.prop.load(input);
    }

    public static Config getInstance() throws IOException {
        if (instance == null) {
            instance = new Config();
        }
        return instance;
    }

    public String getFilePath() {
        return this.prop.getProperty("file_path");
    }

    public String getPassword() {
        return this.prop.getProperty("password");
    }
}
