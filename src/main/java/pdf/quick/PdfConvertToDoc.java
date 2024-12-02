package pdf.quick;


import com.aspose.pdf.Document;
import com.aspose.pdf.SaveFormat;

/**
 * pdf转doc
 *
 * @author pengshuaifeng
 * 2024/12/3
 */
public class PdfConvertToDoc {

    public static void main(String arg[]) {
        // Open the quick PDF document
        Document document = new Document("/Users/pengshuaifeng/IdeaProjects/java-aspose/src/main/resources/quick/市直各类机关名称及简称的通知.pdf");
        // Save the file into MS document format
        document.save("/Users/pengshuaifeng/IdeaProjects/java-aspose/src/main/resources/quick/市直各类机关名称及简称的通知out.doc", SaveFormat.Doc);
        document.close();
    }

}
