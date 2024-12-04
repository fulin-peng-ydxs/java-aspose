package pdf.quick.resolve;

import com.aspose.pdf.Page;
import com.aspose.pdf.PageCollection;
import com.aspose.pdf.XImage;

/**
 * 加载图片
 *
 * @author fulin-peng
 * 2024-12-04  10:45
 */
public class LoadImage {

    public static void main(String[] args) {
        Extract_Images();
    }

    public static void Extract_Images(){
        String _dataDir = "E:\\java-aspose\\src\\main\\resources\\quick\\";
        String filePath = _dataDir + "市直各类机关名称及简称的通知.pdf";
        //加载pdf文件
        com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(filePath);
        //获取所有页面
        PageCollection pages = pdfDocument.getPages();
        int index = 1;
        for (Page page : pages) { //遍历所有页面
            //获取页面中的图片
            com.aspose.pdf.XImageCollection xImageCollection = page.getResources().getImages();
            int imageIndex= 0;
            for (XImage xImage : xImageCollection) { //遍历所有图片
                try {
                    java.io.FileOutputStream outputImage = new java.io.FileOutputStream(_dataDir +index+"-"+imageIndex+"output.jpg");
                    xImage.save(outputImage); //图片输出保存
                    outputImage.close();
                } catch (java.io.FileNotFoundException e) {
                    // TODO: handle exception
                    e.printStackTrace();
                } catch (java.io.IOException e) {
                    // TODO: handle exception
                    e.printStackTrace();
                }
                imageIndex++;
            }
            index++;
        }
    }
}
