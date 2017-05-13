/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ppmaker;

import java.util.ArrayList;

/**
 *
 * @author Freddie
 */
public class MySlide {

    public ArrayList<SlideText> replacementData;
    
    public MySlide(){
        replacementData = new ArrayList<>();
    }
}

class SlideText {

    public String key;
    public String value;

    public SlideText() {
    }

    public SlideText(String key, String value) {
        this.key = key;
        this.value = value;
    }

}
