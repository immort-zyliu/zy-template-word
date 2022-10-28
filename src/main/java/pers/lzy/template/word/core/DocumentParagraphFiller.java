package pers.lzy.template.word.core;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  12:40
 */
public interface DocumentParagraphFiller {

    /**
     * 处理段落
     *
     * @param document  被处理的文档
     * @param paragraph 段落
     * @param paramData 全局参数
     */
    void doProcessParagraph(XWPFDocument document, XWPFParagraph paragraph, Map<String, Object> paramData);

}
