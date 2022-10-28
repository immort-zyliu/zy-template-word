package pers.lzy.template.word.core;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import pers.lzy.template.word.anno.TagOperateHandler;
import pers.lzy.template.word.common.TagParser;
import pers.lzy.template.word.core.handler.OperateParagraphHandler;
import pers.lzy.template.word.exception.OperateWordHandlerInitException;

import java.util.List;
import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/24  13:43
 * 公用判断抽取
 */
public abstract class AbstractOperateParagraphHandler implements OperateParagraphHandler {

    /**
     * 当前 handler 所处理的标签
     */
    protected final String tagName;

    public AbstractOperateParagraphHandler() {

        TagOperateHandler operateHandler = this.getClass().getAnnotation(TagOperateHandler.class);
        if (operateHandler == null) {
            throw new OperateWordHandlerInitException("The OperateParagraphHandler must identify the CellOperateHandler annotation");
        }
        tagName = operateHandler.tagName();
    }


    /**
     * 对 段落 进行个性化处理的方法
     *
     * @param document             要操作的 document
     * @param paragraph            要操作的 paragraph
     * @param params               需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionCalculator 表达式计算器
     */
    @Override
    public void operate(XWPFDocument document, XWPFParagraph paragraph, Map<String, Object> params, ExpressionCalculator expressionCalculator) {

        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null) {
            return;
        }

        // 遍历段落中的每一段文本
        for (XWPFRun run : runs) {
            String expression = run.text();
            // 说明没有解析出来表达式，不需要本handler处理。
            if (StringUtils.isBlank(expression)) {
                return;
            }
            // 说明需要本handler 处理。，则调用目标方法进行处理
            this.doOperate(document, run, params, expression, expressionCalculator);
        }

    }

    /**
     * 对 段落 进行个性化处理的方法
     *
     * @param document             要操作的 document
     * @param run                  要操作的 paragraph 中的run
     * @param params               需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionStr        解析出来的表达式
     * @param expressionCalculator 表达式计算器
     */
    protected abstract void doOperate(XWPFDocument document, XWPFRun run, Map<String, Object> params, String expressionStr, ExpressionCalculator expressionCalculator);
}