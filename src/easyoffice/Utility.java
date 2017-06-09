/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package easyoffice;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Map;

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
            System.err.println("Archivo de datos no encontrado (.\\" + file + ")");
            System.exit(-1);
        }

        return content;
    }

    public static Map<String, List<Object>> getArguments(String[] args) {

        final Map<String, List<Object>> params = new HashMap<>();

        List<Object> options = null;
        for (int i = 0; i < args.length; i++) {
            final String argument = args[i];

            if (argument.charAt(0) == '-') {
                if (argument.length() < 2) {
                    showBadArgumentWarning(argument);
                }

                options = new ArrayList<>();
                params.put(argument.substring(1), options);
            } else if (options != null) {
                options.add(argument);
            } else {
                showBadArgumentWarning(argument);
            }
        }

        return params;
    }

    private static void showBadArgumentWarning(String argument) {
        System.err.println(String.format("Bad argument '%s' ", argument));
        getHelpMessage();
        System.exit(-1);
    }

    public static void showExceptionInfo(boolean verbose, Exception ex) {
        if (verbose) {
            System.err.println("Showing exception extra info:");
            System.err.println(ex.getMessage());
            System.err.println(ex.getLocalizedMessage());
            System.err.println(Arrays.toString(ex.getStackTrace()));
        }
    }

    public static void getHelpMessage() {
        System.out.println("For help: pptmaker -h");
    }

    public static void helpMessage() {
        System.out.println("PPT Maker");
        System.out.println("Syntax: pptmaker [-f template-file] [-d data-file] [-o output-file] [-i template-index]");
        System.out.println();
        System.out.println("Input data file must have the next structure: ");
        System.out.println("- - - - - - - - - - - - - - - - - - - ");
        System.out.println("key=value                           - ");
        System.out.println("other_key=other_value               - ");
        System.out.println("#                                   - ");
        System.out.println("key=next_slide                      - ");
        System.out.println("other_key=the same other slide      - ");
        System.out.println("#                                   - ");
        System.out.println("bla=something else..                - ");
        System.out.println("en so on=boo...                     - ");
        System.out.println("#                                   - ");
        System.out.println("- - - - - - - - - - - - - - - - - - - ");

        System.exit(0);

    }
}
