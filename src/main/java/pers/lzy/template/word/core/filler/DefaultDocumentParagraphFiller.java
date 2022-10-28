package pers.lzy.template.word.core.filler;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pers.lzy.template.word.common.TagParser;
import pers.lzy.template.word.core.DocumentParagraphFiller;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.core.handler.OperateParagraphHandler;

import java.util.Map;

import static pers.lzy.template.word.common.TagParser.formatExpressionInMultiRuns;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  12:47
 */
public class DefaultDocumentParagraphFiller implements DocumentParagraphFiller {

    /**
     * 表达式计算器
     */
    private final ExpressionCalculator expressionCalculator;


    private static final Logger log = LoggerFactory.getLogger(DefaultDocumentParagraphFiller.class);

    /**
     * key: handler's tagName
     * value: OperateParagraphHandler
     */
    private final Map<String, OperateParagraphHandler> operateParagraphHandlerTagMap;

    private static volatile DefaultDocumentParagraphFiller INSTANCE;

    public static DefaultDocumentParagraphFiller getInstance(Map<String, OperateParagraphHandler> operateParagraphHandlerTagMap, ExpressionCalculator expressionCalculator) {

        if (INSTANCE == null) {
            synchronized (DefaultDocumentParagraphFiller.class) {
                if (INSTANCE == null) {
                    log.info("init DefaultDocumentParagraphFiller");
                    INSTANCE = new DefaultDocumentParagraphFiller(operateParagraphHandlerTagMap, expressionCalculator);
                }
            }
        }

        return INSTANCE;

    }

    private DefaultDocumentParagraphFiller(Map<String, OperateParagraphHandler> operateParagraphHandlerTagMap, ExpressionCalculator expressionCalculator) {
        this.operateParagraphHandlerTagMap = operateParagraphHandlerTagMap;
        this.expressionCalculator = expressionCalculator;
    }

    /**
     * 处理段落
     *
     * @param document  被处理的文档
     * @param paragraph 段落
     * @param paramData 全局参数
     */
    @Override
    public void doProcessParagraph(XWPFDocument document, XWPFParagraph paragraph, Map<String, Object> paramData) {

        if (paragraph == null) {
            return;
        }
        String paragraphText = paragraph.getText();
        if (StringUtils.isNotEmpty(paragraphText)) {
            // 获取content中的标签
            String tagName = TagParser.findContentTag(paragraphText);
            if (tagName == null) {
                return;
            }

            // 通过tag找到handler
            OperateParagraphHandler operateParagraphHandler = this.operateParagraphHandlerTagMap.get(tagName);
            if (operateParagraphHandler == null) {
                log.warn("No tag({}) handler found, skipped", tagName);
                return;
            }

            // 去掉 段落中的 tagName (tag 标识)
            TagParser.removeTagName(paragraph, tagName);


            // 格式化 碎片 run
            formatExpressionInMultiRuns(paragraph);


            // 调用handler对cell进行处理
            operateParagraphHandler.operate(document, paragraph, paramData, this.expressionCalculator);

        }

    }
}
