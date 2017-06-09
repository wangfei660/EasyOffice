/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package easyoffice.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import easyoffice.Utility;

/**
 *
 * @author Freddie
 */
public class WordMaker {

    private static final String DEFAULT_FILE_DIR = "./data/";
    private static final String INPUT_PPT_NAME = "notas.docx";
    private static String OUTPUT_PPT_NAME = "output.docx";

    private static XWPFDocument openDoc() {

        XWPFDocument doc = null;
        String path = DEFAULT_FILE_DIR + INPUT_PPT_NAME;

        try {
            File file = new File(path);
            FileInputStream inputstream = new FileInputStream(file);
            doc = new XWPFDocument(inputstream);
        } catch (IOException ex) {
            System.err.println("el documento por defecto no existe (./" + path + ")");
            System.exit(-1);
        }
        return doc;
    }

    public static void main(String[] args) {

        // line command params
        Map<String, List<Object>> params = Utility.getArguments(args);

        // storage for title and content
        // for this version only will content a single data
        HashMap<String, String> data = new HashMap<>();

        // reading template
        XWPFDocument doc = openDoc();

        String title = "Titulo";// params.get("-i").toString();
        String content = "Contenido"; //params.get("-c").toString();

        data.put("title", title);
        data.put("content", content);

        replaceText(doc, data);

        savePpt(doc);
    }

    private static void replaceText(XWPFDocument doc, HashMap<String, String> data) {

        Set<String> keySet = data.keySet();

        for (String key : keySet) {
            for (XWPFParagraph p : doc.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();

                for (XWPFRun run : runs) {
                    if (run.toString().toLowerCase().equals(key)) {
                        run.setText(data.get(key), 0);
                    }
                }
            }
        }
    }

    private static void savePpt(XWPFDocument doc) {

        if (!OUTPUT_PPT_NAME.contains(".docx")) {
            OUTPUT_PPT_NAME += ".docx";
        }

        try {
            File file = new File(OUTPUT_PPT_NAME);
            FileOutputStream out = new FileOutputStream(file);
            doc.write(out);
            out.close();
            System.out.println("Documento guardado");
        } catch (IOException ex) {
            System.err.println("El documento no puede ser guardado");
            System.err.println(ex.getMessage());
        }
    }

}
