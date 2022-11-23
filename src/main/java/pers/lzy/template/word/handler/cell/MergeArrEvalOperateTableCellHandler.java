package pers.lzy.template.word.handler.cell;

import com.google.auto.service.AutoService;
import pers.lzy.template.word.anno.HandlerOrder;
import pers.lzy.template.word.anno.TagOperateHandler;
import pers.lzy.template.word.constant.TagNameConstant;
import pers.lzy.template.word.core.handler.OperateTableCellHandler;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/11/22  17:39
 */
@HandlerOrder(20000)
@TagOperateHandler(tagName = TagNameConstant.M_ARR_TAG_NAME)
@AutoService(OperateTableCellHandler.class)
public class MergeArrEvalOperateTableCellHandler extends AbstractArrEvalOperateTableCellHandler {
    /**
     * 是否需要合并相同单元格
     *
     * @return true:合并，false不合并
     */
    @Override
    protected boolean mergeArrCell() {
        return true;
    }
}
