/*
    Messages used in the app
 */
package easyoffice.util;

/**
 *
 * @author admin
 */
public class Strings {
    
    public static String PPT_NOT_FOUND(String file){
        return String.format("La presentacion por defecto no existe \n (./%s) ", file);
    }
    
    public static String MALFORMED_DATA_FILE(){
        return "Archivo de datos con estructura no adecuada";
    }
    
    public static String CANT_SAVE(){
        return "La Presentacion no puede ser guardada";
    }
    
    public static String NOTHING_REPLACED(int index){
        return String.format("Nada remplazado en slide %d "
                + "Desea especificar otro indice? \"pptmaker [-i indice]\", index");
    }
}
        