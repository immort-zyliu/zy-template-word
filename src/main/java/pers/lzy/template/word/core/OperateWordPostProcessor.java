package pers.lzy.template.word.core;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/11/22  17:51
 * word 填充完成之后，统一的后置处理规范
 */
public interface OperateWordPostProcessor {

    /**
     * 对 填充完成之后 的 word 进行个性化处理的方法
     *
     * @param document             要操作的 document
     * @param params               需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionCalculator 表达式计算器
     */
    void operatePostProcess(XWPFDocument document, Map<String, Object> params, ExpressionCalculator expressionCalculator);
}
