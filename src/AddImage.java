import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

public class AddImage {
    int width;
    int height;
    String imgFile;
    int imgFormat;
    File file;
    AddImage(File file) throws IOException {
        this.file = file;
        BufferedImage bufferedImage = ImageIO.read(file);
        setWidth(bufferedImage.getWidth(), bufferedImage.getHeight());
        imgFile = file.getName();
        imgFormat = getImageFormat(imgFile);
    }
    private void setWidth(int width, int height) {
        double x=width, y=height, z=468, i = z/width;
        if (i<1) {
            x = z;
            y = height*i;
        }
        setHeight(x, y);
    }
    private void setHeight(double width, double height) {
        double x=width, y=height, z=607, i = z/height;
        if (i<1) {
            x = width*i;
            y = z;
        }
        this.width = (int) x;
        this.height = (int) y;
    }

    public void setImageToXwpfRun(XWPFRun xwpfRun) throws IOException, InvalidFormatException {
        xwpfRun.setStyle("Style52");
        xwpfRun.addPicture(Files.newInputStream(file.toPath()), imgFormat, imgFile, Units.toEMU(width), Units.toEMU(height));
    }

    private int getImageFormat(String imgFileName) {
        int format;
        if (imgFileName.endsWith(".emf"))
            format = XWPFDocument.PICTURE_TYPE_EMF;
        else if (imgFileName.endsWith(".wmf"))
            format = XWPFDocument.PICTURE_TYPE_WMF;
        else if (imgFileName.endsWith(".pict"))
            format = XWPFDocument.PICTURE_TYPE_PICT;
        else if (imgFileName.endsWith(".jpeg") || imgFileName.endsWith(".jpg"))
            format = XWPFDocument.PICTURE_TYPE_JPEG;
        else if (imgFileName.endsWith(".png"))
            format = XWPFDocument.PICTURE_TYPE_PNG;
        else if (imgFileName.endsWith(".dib"))
            format = XWPFDocument.PICTURE_TYPE_DIB;
        else if (imgFileName.endsWith(".gif"))
            format = XWPFDocument.PICTURE_TYPE_GIF;
        else if (imgFileName.endsWith(".tiff"))
            format = XWPFDocument.PICTURE_TYPE_TIFF;
        else if (imgFileName.endsWith(".eps"))
            format = XWPFDocument.PICTURE_TYPE_EPS;
        else if (imgFileName.endsWith(".bmp"))
            format = XWPFDocument.PICTURE_TYPE_BMP;
        else {
            return 0;
        }
        return format;
    }
    public static String imageFormatToSuffix(int i) {
        String suffix;

        if (i == XWPFDocument.PICTURE_TYPE_EMF)
            suffix = "emf";
        else if (i == XWPFDocument.PICTURE_TYPE_WMF)
            suffix = "wmf";
        else if (i == XWPFDocument.PICTURE_TYPE_PICT)
            suffix = "pict";
        else if (i == XWPFDocument.PICTURE_TYPE_JPEG)
            suffix = "jpg";
        else if (i == XWPFDocument.PICTURE_TYPE_PNG)
            suffix = "png";
        else if (i == XWPFDocument.PICTURE_TYPE_DIB)
            suffix = "dib";
        else if (i == XWPFDocument.PICTURE_TYPE_GIF)
            suffix = "gif";
        else if (i == XWPFDocument.PICTURE_TYPE_TIFF)
            suffix = "tiff";
        else if (i == XWPFDocument.PICTURE_TYPE_EPS)
            suffix = "eps";
        else if (i == XWPFDocument.PICTURE_TYPE_BMP)
            suffix = "bmp";
        else {
            return "";
        }
        return suffix;
    }
    public int getWidth() {
        return width;
    }
    public int getHeight() {
        return height;
    }
    public int getImgFormat() {
        return imgFormat;
    }
    public String getImgFile() {
        return imgFile;
    }
}
