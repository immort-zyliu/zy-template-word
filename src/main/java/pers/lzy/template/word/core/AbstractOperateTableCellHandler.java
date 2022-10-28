package pers.lzy.template.word.core;

import pers.lzy.template.word.anno.TagOperateHandler;
import pers.lzy.template.word.core.handler.OperateTableCellHandler;
import pers.lzy.template.word.exception.OperateWordHandlerInitException;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  13:35
 */
public abstract class AbstractOperateTableCellHandler implements OperateTableCellHandler {

    /**
     * 当前 handler 所处理的标签
     */
    protected final String tagName;

    public AbstractOperateTableCellHandler() {
        TagOperateHandler operateHandler = this.getClass().getAnnotation(TagOperateHandler.class);
        if (operateHandler == null) {
            throw new OperateWordHandlerInitException("The OperateParagraphHandler must identify the CellOperateHandler annotation");
        }
        tagName = operateHandler.tagName();
    }


}
