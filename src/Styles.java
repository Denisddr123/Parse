import org.apache.poi.ooxml.POIXMLTypeLoader;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.StylesDocument;

import java.io.FileInputStream;
import java.io.IOException;

public class Styles {
    private final CTStyles ctStyles;
    Styles (String fileName) throws IOException, XmlException {
        FileInputStream fileInputStream = new FileInputStream(fileName);
        StylesDocument stylesDoc = StylesDocument.Factory.parse(fileInputStream, POIXMLTypeLoader.DEFAULT_XML_OPTIONS);
        this.ctStyles = stylesDoc.getStyles();
        fileInputStream.close();
    }
    public CTStyles getCtStyles() {
        return ctStyles;
    }
}
