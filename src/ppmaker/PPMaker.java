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
import java.util.List;
import java.util.Map;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 *
 * @author Freddie
 */
public class PPMaker {

    final static String DATA_FILE_NAME = "data.txt";
    final static String STRING_FORMAT = "@%s@";
    final static int TEMPLATE_SLIDE_INDEX = 0;

    static String INPUT_PPT_NAME = "template.pptx";
    static String OUTPUT_PPT_NAME = "output.pptx";

    private static XMLSlideShow openPpt() {
        XMLSlideShow ppt = null;

        try {
            File file = new File(INPUT_PPT_NAME);
            FileInputStream inputstream = new FileInputStream(file);
            ppt = new XMLSlideShow(inputstream);
        } catch (IOException ex) {
            System.err.println("La presentacion seleccionada no existe");
            System.exit(-1);
        }

        return ppt;
    }

    private static void savePpt(XMLSlideShow ppt) {

        if (!OUTPUT_PPT_NAME.contains(".pptx")) {
            OUTPUT_PPT_NAME += ".pptx";
        }

        try {
            File file = new File(OUTPUT_PPT_NAME);
            FileOutputStream out = new FileOutputStream(file);
            ppt.write(out);
            out.close();
            System.out.println("La Presentacion guardada");
        } catch (IOException ex) {
            System.err.println("La Presentacion no puede ser guardada");
            System.err.println(ex.getMessage());
        }
    }

    private static void replaceContent(XSLFTextShape shape, SlideText data) {
        String textInShape = shape.getText();
        String newText = textInShape.replace(String.format(STRING_FORMAT, data.key), data.value);
        shape.setText(newText);
    }

    private static XSLFSlide copySlide(XMLSlideShow ppt, XSLFSlide srcSlide) {
        XSLFSlideLayout layout = srcSlide.getSlideLayout();
        XSLFSlide newSlide = ppt.createSlide(layout);

        return newSlide.importContent(srcSlide);
    }

    private static ArrayList<MySlide> getDataFileContent() {
        String content = Utility.readFile(DATA_FILE_NAME);
        String[] contentLines = content.split("\n");
        ArrayList<MySlide> slidesData = new ArrayList<>();
        MySlide mySlide = new MySlide();

        for (String contentLine : contentLines) {

            if (contentLine.equals("\r")) {
                slidesData.add(mySlide);
                mySlide = new MySlide();
                continue;
            }

            String[] line = contentLine.split("=");

            try {
                mySlide.replacementData.add(new SlideText(line[0], line[1]));
            } catch (IndexOutOfBoundsException ex) {
                System.err.println("Malformed data file");
                System.exit(-1);
            }
        }
        return slidesData;
    }

    public static void main(String[] args) throws IOException {

        Map<String, List<String>> params = Utility.getArguments(args);

        OUTPUT_PPT_NAME = params.getOrDefault("o", new ArrayList<String>() {
            {
                add(OUTPUT_PPT_NAME);
            }
        }).get(0);

        INPUT_PPT_NAME = params.getOrDefault("f", new ArrayList<String>() {
            {
                add(INPUT_PPT_NAME);
            }
        }).get(0);

        XMLSlideShow ppt = openPpt();
        ArrayList<MySlide> replacementData = getDataFileContent();
        XSLFSlide templateSlide = ppt.getSlides().get(TEMPLATE_SLIDE_INDEX);

        replacementData.forEach((slide) -> {
            XSLFSlide copiedSlide = copySlide(ppt, templateSlide);
            XSLFTextShape[] slidePlaceholders = copiedSlide.getPlaceholders();

            for (XSLFTextShape slidePlaceholder : slidePlaceholders) {
                for (SlideText slideText : slide.replacementData) {
                    if (slidePlaceholder.getText().equals(String.format(STRING_FORMAT, slideText.key))) {
                        replaceContent(slidePlaceholder, slideText);
                    }
                }
            }
        });

        ppt.removeSlide(TEMPLATE_SLIDE_INDEX);

        savePpt(ppt);
    }
}
