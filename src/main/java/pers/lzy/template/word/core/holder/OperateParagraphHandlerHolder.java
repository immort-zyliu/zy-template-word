package pers.lzy.template.word.core.holder;

import pers.lzy.template.word.core.handler.OperateParagraphHandler;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/25  13:17
 */
public class OperateParagraphHandlerHolder {

    /**
     * handler
     */
    private final OperateParagraphHandler operateParagraphHandler;

    /**
     * tag name
     */
    private final String handlerTagName;

    public OperateParagraphHandlerHolder(OperateParagraphHandler operateParagraphHandler, String handlerTagName) {
        this.operateParagraphHandler = operateParagraphHandler;
        this.handlerTagName = handlerTagName;
    }

    public OperateParagraphHandler getOperateParagraphHandler() {
        return operateParagraphHandler;
    }

    public String getHandlerTagName() {
        return handlerTagName;
    }
}
