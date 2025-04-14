package words;


import com.aspose.words.*;
import lombok.Builder;
import lombok.Data;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Map;
import java.util.List;

/**
 * Word工具类
 *
 * @author pengshuaifeng
 * 2025/4/14
 */
public class WordsUtils {

    //变量占位符
    public static final String VARIABLE_PREFIX_PLACEHOLDER = "{";
    public static final String VARIABLE_SUFFIX_PLACEHOLDER = "}";

    /**
     * 生成Word文档
     * 2025/4/14 11:36
     * @author pengshuaifeng
     * @param document Word文档对象
     * @param templateSettings 模版设置
     * @return 生成的Word文档字节数组
     */
    public static byte[]  save(Document document, TemplateSettings templateSettings) throws Exception {
        //替换文本变量
        if (templateSettings.textMap!=null) {
            replaceTextVariables(document, templateSettings);
        }
        if (templateSettings.getImageMap()!=null) {
            replaceImageVariables(document, templateSettings);
        }
        //替换表格变量
        if (templateSettings.tableDataMap!=null) {
            replaceTableRows(document, templateSettings);
        }
        //保存Word文档
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        document.save(out,templateSettings.saveFormat);
        return out.toByteArray();
    }

    /**
     * 表格填充
     * 2025/4/14 15:20
     * @author pengshuaifeng
     * @param doc Word文档对象
     * @param templateSettings 模版设置
     */
    public static void replaceTableRows(Document doc,TemplateSettings templateSettings) throws Exception {
        String variablePrefix = getVariablePrefix(templateSettings);
        String variableSuffix = getVariableSuffix(templateSettings);
        for (TableSettings tableSettings : templateSettings.tableDataMap) {
            // 获取指定索引的表格
            Table table = (Table) doc.getChild(NodeType.TABLE, tableSettings.tableIndex, true);
            if (table == null)
                return;
            Row templateRow = null;

            // 找到模板行
            for (Row row : table.getRows()) {
                if (row.toString(SaveFormat.TEXT).contains(VARIABLE_PREFIX_PLACEHOLDER)) {
                    templateRow = row;
                    break;
                }
            }
            if (templateRow == null)
                return;

            //循环遍历数据列表，填充数据
            int rowIndex = table.getRows().indexOf(templateRow);
            for (Map<String, Object> rowData : tableSettings.dataList) {
                Row newRow = (Row) templateRow.deepClone(true);
                for (Cell cell : newRow.getCells()) {  // 遍历单元格：每个单元格循环替换所有占位符
                    for (Map.Entry<String, Object> entry : rowData.entrySet()) {
                        Object value = entry.getValue();
                        if (value instanceof  ImageSettings){
                            replaceImageVariables(doc, entry.getKey(), ((ImageSettings) value).imageStream, variablePrefix,variableSuffix);
                        }else{
                            cell.getRange().replace(variablePrefix + entry.getKey() + variableSuffix, value.toString());
                        }
                    }
                }
                rowIndex=rowIndex+1;
                table.getRows().insert(rowIndex , newRow);
            }
            templateRow.remove(); // 删除模板行
        }
    }


    /**
     * 文本填充
     * 2025/4/14 15:21
     * @author pengshuaifeng
     * @param doc Word文档对象
     * @param templateSettings 模版设置
     */
    public static void replaceTextVariables(Document doc, TemplateSettings templateSettings) throws Exception {
        for (Map.Entry<String, Object> entry : templateSettings.textMap.entrySet()) {
            doc.getRange().replace(getVariablePrefix(templateSettings) + entry.getKey() + getVariableSuffix(templateSettings), entry.getValue().toString());
        }
    }

    /**
     * 图片填充
     * 2025/4/14 15:20
     * @author pengshuaifeng
     * @param doc Word文档对象
     * @param templateSettings 模版设置
     */
    public static void replaceImageVariables(Document doc,TemplateSettings templateSettings)throws Exception {
        for (Map.Entry<String, ImageSettings> entry : templateSettings.imageMap.entrySet()) {
            String key = entry.getKey();
            InputStream imageStream = entry.getValue().imageStream;
            replaceImageVariables(doc, key, imageStream, getVariablePrefix(templateSettings), getVariableSuffix(templateSettings));
        }
    }


    private static void replaceImageVariables(Document doc,String key ,InputStream imageStream,String variablePrefix, String variableSuffix) throws Exception {
        DocumentBuilder builder = new DocumentBuilder(doc);
        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(e -> {
            builder.moveTo(e.getMatchNode());
            builder.insertImage(imageStream);
            e.getMatchNode().remove(); // 移除原变量标记
            return ReplaceAction.SKIP;
        });
        doc.getRange().replace(variablePrefix + key + variableSuffix, "", options);
    }




    /**
     * 获取模版替换变量的前缀占位符
     * 2025/4/14 18:36
     * @author pengshuaifeng
     * @param   templateSettings 模版设置
     */
    public static String getVariablePrefix(TemplateSettings templateSettings) {
        return templateSettings != null ? templateSettings.VARIABLE_PREFIX : VARIABLE_PREFIX_PLACEHOLDER;
    }
    /**
     * 获取模版替换变量的后缀占位符
     * 2025/4/14 18:36
     * @author pengshuaifeng
     * @param   templateSettings 模版设置
     */
    public static String getVariableSuffix(TemplateSettings templateSettings) {
        return templateSettings != null ? templateSettings.VARIABLE_SUFFIX : VARIABLE_SUFFIX_PLACEHOLDER;
    }

    @Data
    public static final  class  TemplateSettings {

        private int saveFormat = SaveFormat.PDF;  //保存格式

        private String VARIABLE_PREFIX = "{"; //变量占位符前缀
        private String VARIABLE_SUFFIX = "}"; //变量占位符后缀

        private int imgWidth = 0; //图片宽度
        private int imgHeight = 0; //图片高度

        //表格数据
        private List<TableSettings> tableDataMap;
        //图片数据
        private Map<String, ImageSettings> imageMap;
        //文本数据
        private Map<String, Object> textMap;
    }

    @Data
    @Builder
    public static final class  TableSettings {

        private int tableIndex; //表格索引

        private List<Map<String, Object>> dataList; //表格数据列表
    }

    @Data
    @Builder
    public static final class ImageSettings {
        private InputStream imageStream; //图片流
    }
}
