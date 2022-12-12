import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.net.URL;
import java.util.ArrayList;

public class NewDocument {
    XWPFDocument document = new XWPFDocument();
    XWPFParagraph xwpfParagraph;
    XWPFRun xwpfRun;
    String fileName = "create.docx";
    int numId = -1;
    ArrayList<Integer> numIdOriginal = new ArrayList<>();
    XWPFNumbering xwpfNumbering;

    NewDocument () throws IOException, XmlException {
        XWPFStyles xwpfStyles = document.createStyles();
        Styles styles = new Styles("styles.xml");
        xwpfStyles.setStyles(styles.getCtStyles());
        XWPFNumbering xwpfNumbering = document.createNumbering();
        new Numbering(xwpfNumbering, "numbering.xml");
        this.xwpfNumbering = xwpfNumbering;
    }
    public void addH1(String str) {
        xwpfParagraph = document.createParagraph();
        xwpfParagraph.setStyle("13");
        xwpfParagraph.setPageBreak(false);
        xwpfRun = xwpfParagraph.createRun();
        xwpfRun.setText(str);
    }
    public void addH2(String str) {
        xwpfParagraph = document.createParagraph();
        xwpfParagraph.setStyle("Style51");
        xwpfRun = xwpfParagraph.createRun();
        xwpfRun.setText(str);
    }
    public void addH3(String str) {
        xwpfParagraph = document.createParagraph();
        xwpfParagraph.setStyle("Style52");
        xwpfParagraph.setNumILvl(BigInteger.valueOf(1));
        xwpfRun = xwpfParagraph.createRun();
        xwpfRun.setText(str);
    }

    private void addNum(int x, int deepLvl) {
        if (numIdOriginal.contains(x)) {
            XWPFNum xwpfNum = this.xwpfNumbering.getNum(BigInteger.valueOf(x));
            BigInteger bigInteger = xwpfNum.getCTNum().getAbstractNumId().getVal();
            int y = numIdOriginal.size() + 21;
            this.xwpfNumbering.addNum(bigInteger, BigInteger.valueOf(y));
            CTNum ctNum = this.xwpfNumbering.getNum(BigInteger.valueOf(y)).getCTNum();
            CTNumLvl ctNumLvl = ctNum.addNewLvlOverride();
            ctNumLvl.setIlvl(BigInteger.valueOf(deepLvl));
            ctNumLvl.addNewStartOverride().setVal(BigInteger.valueOf(1));
            this.numId = y;
        } else {
            this.numId = x;
        }
    }
    public void addEnumeration (ArrayList<String> arrayList, int x, int deepLvl, boolean newList) {

        if (this.numId == -1) {
            this.numId = x;

        }
        if (newList) {
            /*try {
                this.numId = addNum(x, deepLvl);
            } catch (RuntimeException exception) {
                System.out.println(exception);
            }*/
            addNum(x, deepLvl);

            xwpfParagraph = document.createParagraph();
            xwpfParagraph.setStyle("Style47");
            xwpfParagraph.getCTPPr().addNewJc().setVal(STJc.Enum.forInt(1));
            xwpfParagraph.getCTPPr().addNewBidi().setVal("0");
            xwpfParagraph.getCTPPr().addNewRPr();

            xwpfParagraph.setNumILvl(BigInteger.valueOf(deepLvl));
            xwpfParagraph.setNumID(BigInteger.valueOf(this.numId));
            xwpfRun = xwpfParagraph.createRun();
            xwpfRun.getCTR().addNewRPr();

            String str = arrayList.get(0);
            arrayList.remove(0);
            if (str.charAt(0)==8212) {
                str = str.substring(1);
            }
            xwpfRun.setText(str);
        }

        for (String element: arrayList) {
            xwpfParagraph = document.createParagraph();
            xwpfParagraph.setStyle("Style47");
            xwpfParagraph.getCTPPr().addNewJc().setVal(STJc.Enum.forInt(1));
            xwpfParagraph.getCTPPr().addNewBidi().setVal("0");
            xwpfParagraph.getCTPPr().addNewRPr();

            xwpfParagraph.setNumILvl(BigInteger.valueOf(deepLvl));
            xwpfParagraph.setNumID(BigInteger.valueOf(x));
            xwpfRun = xwpfParagraph.createRun();
            xwpfRun.getCTR().addNewRPr();
            if (element.charAt(0)==8212) {
                element = element.substring(1);
            }
            xwpfRun.setText(element);
        }
        this.numIdOriginal.add(this.numId);
    }
    public void addEnumeration (ArrayList<String> arrayList, int x) {

        for (String element: arrayList) {
            xwpfParagraph = document.createParagraph();
            xwpfParagraph.setStyle("Style47");
            xwpfParagraph.setNumID(BigInteger.valueOf(x));
            xwpfRun = xwpfParagraph.createRun();
            if (element.charAt(0)==8212) {
                element = element.substring(1);
            }
            xwpfRun.setText(element);
        }

    }
    public void addImage(String str) throws IOException {
        String subName = str.substring(str.length()-12);
        URL url = new URL(str);
        BufferedImage image1 = ImageIO.read(url);
        File output = new File("images/"+subName);

        ImageIO.write(image1, AddImage.imageFormatToSuffix(image1.getType()), output);
        AddImage images;
        try {
            images = new AddImage(output);
            xwpfParagraph = document.createParagraph();
            xwpfParagraph.setStyle("Style54");
            xwpfRun = xwpfParagraph.createRun();
            images.setImageToXwpfRun(xwpfRun);
        } catch (IOException | InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }
    public void addText(String str) {
        xwpfParagraph = document.createParagraph();
        xwpfParagraph.setStyle("Style47");
        xwpfRun = xwpfParagraph.createRun();
        xwpfRun.setText(str);
    }
    public void addQuote(String str) {
        xwpfParagraph = document.createParagraph();
        xwpfParagraph.setStyle("Style47");
        xwpfRun = xwpfParagraph.createRun();
        xwpfRun.setBold(true);
        xwpfRun.setColor("13112c");
        xwpfRun.setItalic(true);
        xwpfRun.setText(str);
    }
    public void addArticle(String image, String text) throws IOException {
        addImage(image);
        xwpfRun.addBreak();
        xwpfRun.setText(text);
    }
    public void addTable(String[][] strings) {
        XWPFTable xwpfTable = document.createTable(strings.length, strings[0].length);
        xwpfTable.removeBorders();
        String str = "Style59";
        for (int i=0; i<strings.length; i++) {
            XWPFTableRow xwpfTableRow = xwpfTable.getRow(i);
            for (int j=0; j<strings[0].length; j++) {
                XWPFTableCell xwpfTableCell = xwpfTableRow.getCell(j);
                XWPFParagraph paragraph = xwpfTableCell.addParagraph();
                paragraph.setStyle(str);
                paragraph.createRun().setText(strings[i][j]);
            }
            str = "Style55";
        }

    }
    public void addHyperlinkRun(String uri, String text) {
        XWPFHyperlinkRun xwpfHyperlinkRun = xwpfParagraph.createHyperlinkRun(uri);
        xwpfHyperlinkRun.setText(text);
        xwpfHyperlinkRun.setColor("0000FF");
    }
    public void writeDocument() throws IOException {
        FileOutputStream out = new FileOutputStream(this.fileName);
        document.write(out);
        out.close();
        document.close();
    }
    public void writeDocument(String str) throws IOException {
        FileOutputStream out = new FileOutputStream(str);
        document.write(out);
        out.close();
        document.close();
    }
}
