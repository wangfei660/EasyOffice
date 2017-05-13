/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ppmaker;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 *
 * @author Freddie
 */
public class Utility {

    public static String readFile(String file) {
        String content = null;
        try {
            content = new String(Files.readAllBytes(Paths.get(file)));
        } catch (IOException ex) {
            System.err.println("Archivo de datos no encontrado");
            System.err.println(ex.getMessage());
            System.exit(-1);
        }
        
        return content;
    }
}
