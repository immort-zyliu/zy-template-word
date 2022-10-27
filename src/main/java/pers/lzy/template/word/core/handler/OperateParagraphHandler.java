package pers.lzy.template.word.core.handler;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import pers.lzy.template.word.core.ExpressionCalculator;

import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/27  16:39
 * 操作段落的handler
 */
public interface OperateParagraphHandler {

    /**
     * 对 段落 进行个性化处理的方法
     *
     * @param document             要操作的 document
     * @param run                  要操作的 paragraph 中的run
     * @param params               需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionCalculator 表达式计算器
     */
    void operate(XWPFDocument document, XWPFRun run, Map<String, Object> params, ExpressionCalculator expressionCalculator);
}
