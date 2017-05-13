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
import org.apache.poi.xslf.usermodel.SlideLayout;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 *
 * @author Freddie
 */
public class PPMaker {

    static void edit() throws IOException {

        //opening an existing slide show
        File file = new File("example1.pptx");
        FileInputStream inputstream = new FileInputStream(file);
        XMLSlideShow ppt = new XMLSlideShow(inputstream);

        //adding slides to the slodeshow
        XSLFSlide slide1 = ppt.createSlide();
        XSLFSlide slide2 = ppt.createSlide();

        //saving the changes 
        FileOutputStream out = new FileOutputStream(file);
        ppt.write(out);

        System.out.println("Presentation edited successfully");
        out.close();
    }

    public static void create() throws IOException {
        //creating a new empty slide show
        XMLSlideShow ppt = new XMLSlideShow();

        //creating an FileOutputStream object
        File file = new File("example1.pptx");
        FileOutputStream out = new FileOutputStream(file);

        //saving the changes to a file
        ppt.write(out);
        System.out.println("Presentation created successfully");
        out.close();
    }

    public static void prueba() throws IOException {
        File file = new File("example1.pptx");
        FileInputStream inputstream = new FileInputStream(file);
        XMLSlideShow ppt = new XMLSlideShow(inputstream);

        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    System.out.println(textShape.getText());
                }
            }
            try {

                for (XSLFShape shape : slide.getNotes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape txShape = (XSLFTextShape) shape;
                        for (XSLFTextParagraph xslfParagraph : txShape.getTextParagraphs()) {
                            System.out.println(xslfParagraph.getText());
                        }
                    }
                }
            } catch (Exception e) {
            }

        }

//        FileOutputStream out = new FileOutputStream(file);
//        ppt.write(out);
//        out.close();
    }

    public static void main(String[] args) throws IOException {
        prueba();
    }
}
