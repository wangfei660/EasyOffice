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
import java.util.HashMap;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;

/**
 *
 * @author Freddie
 */
public class PPMaker {

    static String EXAMPLE_PPT_NAME = "example1.pptx";

    private static XMLSlideShow openPpt(File pptFile) {
        XMLSlideShow ppt = null;

        try {
            FileInputStream inputstream = new FileInputStream(pptFile);
            ppt = new XMLSlideShow(inputstream);
        } catch (IOException ex) {
            System.out.println("La presentacion seleccionada no existe");
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

    public static XSLFSlide copySlide(XMLSlideShow ppt, XSLFSlide srcSlide){
        XSLFSlideLayout layout = srcSlide.getSlideLayout();
        XSLFSlide newSlide = ppt.createSlide(layout);

        return newSlide.importContent(srcSlide);
    }

    public static void main(String[] args) throws IOException {
        File f = new File(EXAMPLE_PPT_NAME);
        XMLSlideShow ppt = openPpt(f);

        XSLFSlide srcSlide = ppt.getSlides().get(0);
        
        copySlide(ppt, srcSlide);
        copySlide(ppt, srcSlide);
        copySlide(ppt, srcSlide);
        copySlide(ppt, srcSlide);

        savePpt(f, ppt);
    }
}
