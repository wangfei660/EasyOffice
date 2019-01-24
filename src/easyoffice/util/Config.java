/*
    Class for reading config file "config.cfg"
 */
package easyoffice.util;

/**
 *
 * @author admin
 */
import java.util.Properties;

public class Config {

    private Properties configFile;

    private static Config instance;

    private Config() {
        configFile = new java.util.Properties();
        try {
            configFile.load(this.getClass().getClassLoader().getResourceAsStream("config.cfg"));
        } catch (Exception eta) {
            eta.printStackTrace();
        }
    }

    private String getValue(String key) {
        return configFile.getProperty(key);
    }

    public static Object getProperty(String key) {
        if (instance == null) {
            instance = new Config();
        }
        return instance.getValue(key);
    }
}
