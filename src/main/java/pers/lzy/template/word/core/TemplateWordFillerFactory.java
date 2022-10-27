package pers.lzy.template.word.core;

import pers.lzy.template.word.filler.DefaultTemplateWordFiller;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/23  18:25
 */
public class TemplateWordFillerFactory {


    /**
     * 获取默认 excel 导入工具的 构造器
     */
    public static DefaultTemplateWordFiller.Builder defaultTemplateWordFillerBuilder() {
        return new DefaultTemplateWordFiller.Builder();
    }


}
