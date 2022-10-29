package pers.lzy.template.word.handler.cell;

import com.google.auto.service.AutoService;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import pers.lzy.template.word.anno.HandlerOrder;
import pers.lzy.template.word.anno.TagOperateHandler;
import pers.lzy.template.word.common.TagParser;
import pers.lzy.template.word.constant.TagNameConstant;
import pers.lzy.template.word.core.AbstractOperateTableCellHandler;
import pers.lzy.template.word.core.DocumentParagraphFiller;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.core.handler.OperateTableCellHandler;
import pers.lzy.template.word.pojo.WordCell;
import pers.lzy.template.word.pojo.WordRow;
import pers.lzy.template.word.pojo.WordTable;

import java.util.List;
import java.util.Map;

import static pers.lzy.template.word.common.TagParser.formatExpressionInMultiRuns;
import static pers.lzy.template.word.common.TagParser.verifyHasExpression;
import static pers.lzy.template.word.utils.WordUtil.setRunValue;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/29  13:25
 */
@HandlerOrder(20000)
@TagOperateHandler(tagName = TagNameConstant.TABLE_SIMPLE_TAG_NAME)
@AutoService(OperateTableCellHandler.class)
public class TableSimpleEvalOperateTableCellHandler extends AbstractOperateTableCellHandler {


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
    @Override
    public void operate(XWPFDocument document, WordTable table, WordRow row, WordCell cell, Map<String, Object> params, ExpressionCalculator expressionCalculator, DocumentParagraphFiller documentParagraphFiller) {

        List<XWPFParagraph> paragraphs = cell.paragraphs();
        if (paragraphs == null) {
            return;
        }

        // 移除 标签
        TagParser.removeTagName(paragraphs, this.tagName);

        for (XWPFParagraph paragraph : paragraphs) {
            // 整理 runs
            formatExpressionInMultiRuns(paragraph);

            // 操作填充段落
            operateParagraph(document, paragraph, params, expressionCalculator);
        }


    }


    private void operateParagraph(XWPFDocument document, XWPFParagraph paragraph, Map<String, Object> params, ExpressionCalculator expressionCalculator) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null) {
            return;
        }

        // 遍历段落中的每一段文本
        for (XWPFRun run : runs) {
            String expression = run.text();
            // 说明没有解析出来表达式，不需要本handler处理。
            if (!verifyHasExpression(expression)) {
                continue;
            }
            // 说明需要本handler 处理。，则调用目标方法进行处理
            this.doOperate(document, run, params, expression, expressionCalculator);
        }
    }

    private void doOperate(XWPFDocument document, XWPFRun run, Map<String, Object> params, String expression, ExpressionCalculator expressionCalculator) {
        // 说明需要处理, 计算表达式并赋值。
        Object result = expressionCalculator.calculateNoFormat(expression, params);
        // 设置到单元格中
        setRunValue(run, this.formatCellValue(result));
    }

    /**
     * 格式化 计算出来两单元格的值,子类可以重写更改
     *
     * @param realValue 计算出来的值
     * @return 格式化后的值
     */
    protected Object formatCellValue(Object realValue) {
        return realValue;
    }
}
