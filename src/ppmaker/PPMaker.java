/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ppmaker;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;

/**
 *
 * @author Freddie
 */
public class PPMaker {

    final static String EXAMPLE_PPT_NAME = "example1.pptx";
    final static String DATA_FILE_NAME = "data.txt";

    private static XMLSlideShow openPpt(File pptFile) {
        XMLSlideShow ppt = null;

        try {
            FileInputStream inputstream = new FileInputStream(pptFile);
            ppt = new XMLSlideShow(inputstream);
        } catch (IOException ex) {
            System.err.println("La presentacion seleccionada no existe");
            System.exit(-1);
        }

        return ppt;
    }

    private static void savePpt(File file, XMLSlideShow ppt) {
        try {
            FileOutputStream out = new FileOutputStream(file);
            ppt.write(out);
            out.close();
            System.out.println("La Presentacion guardada");
        } catch (IOException ex) {
            System.err.println("La Presentacion no puede ser guardada");
            System.err.println(ex.getMessage());
        }
    }

    private static String replaceContent(String content, HashMap<String, String> data) {
        return content.replace(String.format("@%s@", data.get("key")), data.get("value"));
    }

    private static XSLFSlide copySlide(XMLSlideShow ppt, XSLFSlide srcSlide) {
        XSLFSlideLayout layout = srcSlide.getSlideLayout();
        XSLFSlide newSlide = ppt.createSlide(layout);

        return newSlide.importContent(srcSlide);
    }

    private static ArrayList<DataRow> getDataFileContent() {
        String content = Utility.readFile(DATA_FILE_NAME);
        String[] contentLines = content.split("\n");
        ArrayList<DataRow> dataList = new ArrayList<>();
        
        for (String contentLine : contentLines) {
            String[] line = contentLine.split("=");
            try {
                dataList.add(new DataRow(line[0], line[1]));
            } catch (IndexOutOfBoundsException ex) {
                System.err.println("Malformed data file");
                System.exit(-1);
            }
        }
        return dataList;
    }

    public static void main(String[] args) throws IOException {


        /*
            
        File file = new File(EXAMPLE_PPT_NAME);
        XMLSlideShow ppt = openPpt(file);

        ArrayList<DataRow> data = getDataFileContent();

        XSLFSlide srcSlide = ppt.getSlides().get(0);
            
            for (int i = 0; i < 1; i++) {
            XSLFSlide slide = copySlide(ppt, srcSlide);
            
            XSLFTextShape[] a= slide.getPlaceholders();
            
            System.out.println(a[0].getText());

            XSLFTextShape title2 = slide.getPlaceholder(0);
            title2.addNewTextParagraph().addNewTextRun().setText("BOO");
            title2.addNewTextParagraph().addLineBreak();
            title2.addNewTextParagraph().addNewTextRun().setText("BOO");
            //            title2.setText("Second Title");

            XSLFTextShape body2 = slide.getPlaceholder(1);
            body2.clearText(); // unset any existing text
            body2.addNewTextParagraph().addNewTextRun().setText("First paragraph");
            body2.addNewTextParagraph().addNewTextRun().setText("Second paragraph");
            body2.addNewTextParagraph().addNewTextRun().setText("Third paragraph");
            }
         */
//        savePpt(f, ppt);
    }
}
