/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ppmaker;

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
import org.apache.poi.sl.usermodel.StrokeStyle;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 *
 * @author Freddie
 */
public class PPMaker {

    final static String DATA_FILE_NAME = "data.txt";
    final static String STRING_FORMAT = "@%s@";

    static int TEMPLATE_SLIDE_INDEX = 0;
    static String INPUT_PPT_NAME = "template.pptx";
    static String OUTPUT_PPT_NAME = "output.pptx";
    static boolean VERBOSE = false;

    private static XMLSlideShow openPpt() {
        XMLSlideShow ppt = null;

        try {
            File file = new File(INPUT_PPT_NAME);
            FileInputStream inputstream = new FileInputStream(file);
            ppt = new XMLSlideShow(inputstream);
        } catch (IOException ex) {
            System.err.println("La presentacion por defecto no existe (./" + INPUT_PPT_NAME + ")");
            System.err.println("Para seleccionar \"pptmaker -f [input-file]\"");
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

        XSLFTextRun addedText = shape.setText(newText.replace("\r", "")); // not include break line
        addedText.setFontFamily(HSSFFont.FONT_ARIAL);
        addedText.setFontColor(Color.BLACK);
    }

    private static XSLFSlide copySlide(XMLSlideShow ppt, XSLFSlide srcSlide) {
        XSLFSlideLayout layout = srcSlide.getSlideLayout();
        XSLFSlide newSlide = ppt.createSlide(layout);
        ppt.setSlideOrder(newSlide, srcSlide.getSlideNumber());

        return newSlide.importContent(srcSlide);
    }

    private static ArrayList<MySlide> getDataFileContent() {
        String content = Utility.readFile(DATA_FILE_NAME);
        String[] contentLines = content.split("\n");
        ArrayList<MySlide> slidesData = new ArrayList<>();
        MySlide mySlide = new MySlide();

        for (String contentLine : contentLines) {

            if (contentLine.equals("#\r") || contentLine.equals("#")) {
                slidesData.add(mySlide);
                mySlide = new MySlide();
                continue;
            }

            String[] line = contentLine.split("=");

            try {
                mySlide.replacementData.add(new SlideText(line[0], line[1]));
            } catch (IndexOutOfBoundsException ex) {
                System.err.println("Archivo de datos con estructura no adecuada");
                Utility.helpMessage();
                System.exit(-1);
            }
        }
        return slidesData;
    }

    private static void evaluateArguments(String[] args) {
        Map<String, List<Object>> params = Utility.getArguments(args);

        // check for help
        if (params.containsKey("h")) {
            Utility.helpMessage();
        }

        if (params.containsKey("v")) {
            VERBOSE = true;
        }

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

    private static void setFirstSlideInfo(XMLSlideShow ppt) {
        XSLFSlide firstSlide = ppt.getSlides().get(0);

        MySlide mySlide = new MySlide();
        mySlide.replacementData.add(new SlideText("Nombre_Mes", LocalDate.now().getMonth().getDisplayName(TextStyle.FULL.FULL, new Locale("es", "ES"))));
        mySlide.replacementData.add(new SlideText("Año", String.valueOf(LocalDate.now().getYear())));

        List<XSLFShape> slideShapes = firstSlide.getShapes();

        for (XSLFShape slideShape : slideShapes) {
            if (slideShape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) slideShape;
                for (SlideText slideText : mySlide.replacementData) {
                    if (textShape.getText().contains(String.format(STRING_FORMAT, slideText.key))) {
                        replaceContent(textShape, slideText);
                    }
                }
            }
        }
    }

    private static void setEngineData(XMLSlideShow ppt) {
        ArrayList<MySlide> replacementData = getDataFileContent();
        XSLFSlide templateSlide = ppt.getSlides().get(TEMPLATE_SLIDE_INDEX);
        boolean somethingReplaced = false;

        for (MySlide slide : replacementData) {
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
            System.out.println("Nada remplazado en slide " + TEMPLATE_SLIDE_INDEX +" ");
            System.out.println("Desea especificar otro indice? \"pptmaker [-i indice]\"");
            System.exit(0);
        }
    }

    public static void main(String[] args) throws IOException {

        evaluateArguments(args);

        XMLSlideShow ppt = openPpt();

        setFirstSlideInfo(ppt);
        setEngineData(ppt);
        ppt.removeSlide(TEMPLATE_SLIDE_INDEX); // remove the template slide
        savePpt(ppt);
    }
}
