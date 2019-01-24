/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package easyoffice.powerpoint.models;

/**
 *
 * @author Freddie
 */

public class SlideText {

    public String key;
    public String value;

    public SlideText() {
    }

    public SlideText(String key, String value) {
        this.key = key;
        this.value = value;
    }

    @Override
    public String toString() {
        return String.format("%s=%s", this.key, this.value);
    }
}

