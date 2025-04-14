package words;


import com.aspose.words.Document;

import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author pengshuaifeng
 * 2025/4/15
 */
public class WordsUtilsQuick {

    public static void main(String[] args) throws Exception {
        Document document = new Document(WordsUtils.class.getResourceAsStream("/quick/words/hidden_dangers_notification.docx"));
        WordsUtils.TemplateSettings settings = new WordsUtils.TemplateSettings();
        Map<String, Object> textMap = new HashMap<>();
        textMap.put("company", "张三");
        textMap.put("number", "1、2");
        textMap.put("year", 2025);
        textMap.put("month",4);
        textMap.put("day", 20);
        settings.setTextMap(textMap);
        byte[] bytes = WordsUtils.save(document, settings);
        //bytes保存本地为pdf
        System.out.println("Word文档生成成功，大小：" + bytes.length + "字节");
        //保存为pdf
        FileOutputStream fileOutputStream = new FileOutputStream("/Users/pengshuaifeng/IdeaProjects/java-aspose/src/main/resources/output.pdf");
        fileOutputStream.write(bytes);
        fileOutputStream.close();
    }

}
