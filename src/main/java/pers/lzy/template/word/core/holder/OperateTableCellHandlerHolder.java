package pers.lzy.template.word.core.holder;

import pers.lzy.template.word.core.handler.OperateParagraphHandler;
import pers.lzy.template.word.core.handler.OperateTableCellHandler;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/27  16:58
 */
public class OperateTableCellHandlerHolder {

    /**
     * handler
     */
    private final OperateTableCellHandler operateTableCellHandler;

    /**
     * tag name
     */
    private final String handlerTagName;

    public OperateTableCellHandlerHolder(OperateTableCellHandler operateTableCellHandler, String handlerTagName) {
        this.operateTableCellHandler = operateTableCellHandler;
        this.handlerTagName = handlerTagName;
    }

    public String getHandlerTagName() {
        return handlerTagName;
    }

    public OperateTableCellHandler getOperateTableCellHandler() {
        return operateTableCellHandler;
    }
}
