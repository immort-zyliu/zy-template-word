package pers.lzy.template.word.handler.run;

import com.google.auto.service.AutoService;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import pers.lzy.template.word.anno.HandlerOrder;
import pers.lzy.template.word.anno.TagOperateHandler;
import pers.lzy.template.word.constant.TagNameConstant;
import pers.lzy.template.word.core.AbstractOperateParagraphHandler;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.core.handler.OperateParagraphHandler;

import java.util.Map;

import static pers.lzy.template.word.utils.WordUtil.setRunValue;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/23  18:12
 * 处理 simple 标签的 handler
 */
@HandlerOrder(10000)
@TagOperateHandler(tagName = TagNameConstant.SIMPLE_TAG_NAME)
@AutoService(OperateParagraphHandler.class)
public class ParagraphSimpleEvalOperateHandler extends AbstractOperateParagraphHandler {


    /**
     * 对 段落 进行个性化处理的方法
     *
     * @param document             要操作的 document
     * @param run                  要操作的 paragraph 中的run
     * @param params               需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionStr        解析出来的表达式
     * @param expressionCalculator 表达式计算器
     */
    @Override
    protected void doOperate(XWPFDocument document, XWPFRun run, Map<String, Object> params, String expressionStr, ExpressionCalculator expressionCalculator) {
        // 说明需要处理, 计算表达式并赋值。
        Object result = expressionCalculator.calculateNoFormat(expressionStr, params);
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
