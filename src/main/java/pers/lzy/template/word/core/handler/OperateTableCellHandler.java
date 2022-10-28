package pers.lzy.template.word.core.handler;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pers.lzy.template.word.core.DocumentParagraphFiller;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.pojo.WordCell;
import pers.lzy.template.word.pojo.WordRow;
import pers.lzy.template.word.pojo.WordTable;

import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/27  16:41
 */
public interface OperateTableCellHandler {

    /**
     * 对 word 表格中的 cell 进行个性化处理的方法
     *
     * @param document                被操作的文档
     * @param table                   被操作的表格
     * @param row                     被操作的行
     * @param cell                    被操作的单元格
     * @param params                  需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionCalculator    表达式计算器
     * @param documentParagraphFiller 文档段落填充器
     */
    void operate(XWPFDocument document, WordTable table, WordRow row, WordCell cell, Map<String, Object> params, ExpressionCalculator expressionCalculator, DocumentParagraphFiller documentParagraphFiller);

}
