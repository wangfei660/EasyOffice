/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package easyoffice.powerpoint;

import easyoffice.powerpoint.models.Slide;
import easyoffice.powerpoint.models.SlideText;
import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFFont;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import easyoffice.util.Utility;
import easyoffice.util.Config;
import easyoffice.util.Strings;
import java.util.Arrays;
import java.util.Collections;
import org.apache.poi.util.ArrayUtil;

/**
 *
 * @author Freddie
 */
public class PPTMaker {

    static String DATA_FILE_NAME = (String) Config.getProperty("data_file_name");
    static String SLIDE_SEPARATOR = (String) Config.getProperty("slide_separator");
    static int TEMPLATE_SLIDE_INDEX = Integer.valueOf(Config.getProperty("template_slide_index").toString());
    static String INPUT_PPT_NAME = (String) Config.getProperty("input_ppt_name");
    static String OUTPUT_PPT_NAME = (String) Config.getProperty("output_ppt_name");
    static String STRING_FORMAT = (String) Config.getProperty("string_format");

    static boolean VERBOSE = false;

    private static XMLSlideShow openPpt() {
        XMLSlideShow ppt = null;

        try {
            File file = new File(INPUT_PPT_NAME);
            FileInputStream inputstream = new FileInputStream(file);
            ppt = new XMLSlideShow(inputstream);
        } catch (IOException ex) {
            System.err.println(Strings.PPT_NOT_FOUND(INPUT_PPT_NAME));
            System.exit(-1);
        }

        return ppt;
    }

    private static boolean savePpt(XMLSlideShow ppt) {

        /*
        * Save PPT instance on disk and add pptx extension
        * if not included in filename
         */
        if (!OUTPUT_PPT_NAME.endsWith(".pptx")) {
            OUTPUT_PPT_NAME += ".pptx";
        }

        try {
            File file = new File(OUTPUT_PPT_NAME);
            FileOutputStream out = new FileOutputStream(file);
            ppt.write(out);
            out.close();

            return true;
        } catch (IOException ex) {
            return false;
        }
    }

    private static void replaceContent(XSLFTextShape shape, SlideText data) {
        /*
        Given a slide shape, replace it content with SlideText data
         */
        String textInShape = shape.getText();
        String newText = textInShape.replace(String.format(STRING_FORMAT, data.key), data.value);

        XSLFTextRun addedText = shape.setText(newText.replace("\r", "")); // not include break line
        addedText.setFontFamily(HSSFFont.FONT_ARIAL);
        addedText.setFontColor(Color.BLACK);
    }

    private static XSLFSlide copySlide(XMLSlideShow ppt, XSLFSlide srcSlide) {
        /*
         Create new slide instance copy
        */
        XSLFSlideLayout layout = srcSlide.getSlideLayout();
        XSLFSlide newSlide = ppt.createSlide(layout);
        ppt.setSlideOrder(newSlide, srcSlide.getSlideNumber());

        return newSlide.importContent(srcSlide);
    }

    private static String[] fileDataContent() {
        String content = Utility.readFile(DATA_FILE_NAME);
        String[] contentLines = content.split(SLIDE_SEPARATOR);
        return contentLines;
    }

    private static ArrayList<Slide> generateSlides() {
        // data file content
        String[] slidesContent = fileDataContent();
        
        // idk why ppt generates reversed, this was the solution
        Collections.reverse(Arrays.asList(slidesContent));
        
        // to put proccessed slides
        ArrayList<Slide> slidesData = new ArrayList<>();

        for (String slideContent : slidesContent ) {
            // separate by break lines
            String[] contentLine = Utility.cleanStr(slideContent).split("\n");
            Slide mySlide = new Slide();

            for (String line : contentLine) {

                try {
                    String[] row = line.split("=");
                    String key = row[0];
                    String value = row[1];
                    mySlide.replacementData.add(new SlideText(key, value));
                } catch (IndexOutOfBoundsException ex) {
                    System.err.println(Strings.MALFORMED_DATA_FILE());
                    Utility.getHelpMessage();
                    System.exit(-1);
                }
            }
            slidesData.add(mySlide);
        }
        return slidesData;
    }

    private static void evaluateArguments(String[] args) throws NumberFormatException {
        Map<String, List<Object>> params = null;

        params = Utility.getArguments(args);

        // check for help
        if (params.containsKey("h")) {
            Utility.helpMessage();
        }

        if (params.containsKey("v")) {
            VERBOSE = true;
        }

        DATA_FILE_NAME = params.getOrDefault("d", new ArrayList<Object>() {
            {
                add(DATA_FILE_NAME);
            }
        }).get(0).toString();

        OUTPUT_PPT_NAME = params.getOrDefault("o", new ArrayList<Object>() {
            {
                add(OUTPUT_PPT_NAME);
            }
        }).get(0).toString();

        INPUT_PPT_NAME = params.getOrDefault("f", new ArrayList<Object>() {
            {
                add(INPUT_PPT_NAME);
            }
        }).get(0).toString();

        TEMPLATE_SLIDE_INDEX = Integer.parseInt(params.getOrDefault("i", new ArrayList<Object>() {
            {
                add(TEMPLATE_SLIDE_INDEX);
            }
        }).get(0).toString());
    }

    private static void processSlides(XMLSlideShow ppt) {
        ArrayList<Slide> replacementData = generateSlides();
        XSLFSlide templateSlide = ppt.getSlides().get(TEMPLATE_SLIDE_INDEX);
        boolean somethingReplaced = false;

        for (Slide slide : replacementData) {
            XSLFSlide copiedSlide = copySlide(ppt, templateSlide);
            List<XSLFShape> slideShapes = copiedSlide.getShapes();

            for (XSLFShape slideShape : slideShapes) {
                if (slideShape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) slideShape;
                    for (SlideText slideText : slide.replacementData) {
                        if (textShape.getText().contains(String.format(STRING_FORMAT, slideText.key))) {
                            replaceContent(textShape,
                                    slideText);
                            somethingReplaced = true;
                        }
                    }
                }
            }
        }

        if (!somethingReplaced) {
            System.out.println(Strings.NOTHING_REPLACED(TEMPLATE_SLIDE_INDEX));
            System.exit(0);
        }
    }

    public static void main(String[] args) throws IOException {
        try {
            evaluateArguments(args);
        } catch (NumberFormatException ex) {
            System.err.println("El parametro debe ser numerico");
            Utility.getHelpMessage();
            System.exit(-1);
        }

        // create ppt instance
        XMLSlideShow ppt = openPpt();
        
        // generate new slides based on data file
        processSlides(ppt);

        // remove the template slide
        ppt.removeSlide(TEMPLATE_SLIDE_INDEX);
        
        // finishing
        savePpt(ppt);
    }
}
