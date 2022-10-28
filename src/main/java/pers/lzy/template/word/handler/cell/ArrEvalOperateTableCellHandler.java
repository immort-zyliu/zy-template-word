package pers.lzy.template.word.handler.cell;

import com.google.auto.service.AutoService;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import pers.lzy.template.word.anno.HandlerOrder;
import pers.lzy.template.word.anno.TagOperateHandler;
import pers.lzy.template.word.constant.TagNameConstant;
import pers.lzy.template.word.core.AbstractOperateTableCellHandler;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.core.handler.OperateTableCellHandler;

import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  13:36
 */
@HandlerOrder(20000)
@TagOperateHandler(tagName = TagNameConstant.ARR_TAG_NAME)
@AutoService(OperateTableCellHandler.class)
public class ArrEvalOperateTableCellHandler extends AbstractOperateTableCellHandler {


    /**
     * 对 word 表格中的 cell 进行个性化处理的方法
     *
     * @param document             被操作的文档
     * @param table                被操作的表格
     * @param row                  被操作的行
     * @param cell                 被操作的单元格
     * @param params               需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionCalculator 表达式计算器
     */
    @Override
    public void operate(XWPFDocument document, XWPFTable table, XWPFTableRow row, XWPFTableCell cell, Map<String, Object> params, ExpressionCalculator expressionCalculator) {


    }
}
