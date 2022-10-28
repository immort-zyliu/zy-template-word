package pers.lzy.template.word.filler;

import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JxltEngine;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pers.lzy.template.word.anno.TagOperateHandler;
import pers.lzy.template.word.calculator.Jxel3ExpressionCalculator;
import pers.lzy.template.word.common.TagParser;
import pers.lzy.template.word.constant.CommonDataNameConstant;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.core.TemplateWordFiller;
import pers.lzy.template.word.core.handler.OperateParagraphHandler;
import pers.lzy.template.word.core.handler.OperateTableCellHandler;
import pers.lzy.template.word.core.holder.OperateParagraphHandlerHolder;
import pers.lzy.template.word.core.holder.OperateTableCellHandlerHolder;
import pers.lzy.template.word.exception.OperateWordHandlerInitException;
import pers.lzy.template.word.pojo.ArrInfo;
import pers.lzy.template.word.pojo.MergeArrInfo;
import pers.lzy.template.word.provider.DocumentProvider;
import pers.lzy.template.word.provider.FillDataProvider;
import pers.lzy.template.word.provider.FunctionProvider;
import pers.lzy.template.word.utils.SpiLoader;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static pers.lzy.template.word.common.TagParser.formatExpressionInMultiRuns;
import static pers.lzy.template.word.utils.WordUtil.mergeRunText;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/27  16:54
 */
public class DefaultTemplateWordFiller implements TemplateWordFiller {

    private static final Logger logger = LoggerFactory.getLogger(DefaultTemplateWordFiller.class);

    /**
     * 表达式计算器
     */
    private final ExpressionCalculator expressionCalculator;

    /**
     * key: handler's tagName
     * value: OperateParagraphHandler
     */
    private final Map<String, OperateParagraphHandler> operateParagraphHandlerTagMap;


    /**
     * key: handler's tagName
     * value: OperateTableCellHandler
     */
    private final Map<String, OperateTableCellHandler> operateTableCellHandlerTagMap;


    public DefaultTemplateWordFiller(Builder builder) {
        int expressionCacheSize = builder.expressionCacheSize;

        // 设置表达式计算器
        if (builder.expressionCalculator != null) {
            this.expressionCalculator = builder.expressionCalculator;
        } else {
            this.expressionCalculator = new Jxel3ExpressionCalculator(builder.jxltEngine, expressionCacheSize);
        }

        this.operateParagraphHandlerTagMap = builder.operateParagraphHandlerTagMap;
        this.operateTableCellHandlerTagMap = builder.operateTableCellHandlerTagMap;
    }

    /**
     * 给sheet填充对应的数据，根据模板
     *
     * @param documentProvider doc 提供者
     * @param dataProvider     数据提供着
     */
    @Override
    public void fillData(DocumentProvider documentProvider, FillDataProvider dataProvider) {
        XWPFDocument document = documentProvider.getDocument();
        Map<String, Object> paramData = dataProvider.getParamData();

        if (document == null || paramData == null) {
            throw new IllegalStateException("document and paramData cannot be null");
        }

        verifyAndInitParamData(paramData);

        processParagraphs(document, paramData);

        processTables(document, paramData);
    }

    private void processTables(XWPFDocument document, Map<String, Object> paramData) {
        logger.info("process table");
        List<XWPFTable> tables = document.getTables();//获取表格

        if (tables == null) {
            return;
        }

        for (XWPFTable table : tables) {
            List<XWPFTableRow> rows = table.getRows();
            if (rows == null) {
                continue;
            }

            for (XWPFTableRow row : rows) {
                List<XWPFTableCell> tableCells = row.getTableCells();
                if (tableCells == null) {
                    return;
                }

                for (XWPFTableCell cell : tableCells) {
                    if (cell != null) {
                        doProcessCell(document, table, row, cell, paramData);
                    }
                }
            }
        }

    }

    private void doProcessCell(XWPFDocument document, XWPFTable table, XWPFTableRow row, XWPFTableCell cell, Map<String, Object> paramData) {

        // 获取cell中的内容
        String content = cell.getText();
        if (StringUtils.isNotBlank(content)) {
            // 获取content中的标签
            String tagName = TagParser.findFirstTag(content);
            if (tagName == null) {
                return;
            }

            // 通过tag找到handler
            OperateTableCellHandler operateTableCellHandler = this.operateTableCellHandlerTagMap.get(tagName);
            if (operateTableCellHandler == null) {
                logger.warn("No tag({}) handler found, skipped", tagName);
                return;
            }

            // 调用handler对cell进行处理
            operateTableCellHandler.operate(document, table, row, cell, paramData, this.expressionCalculator);
        }

    }

    private void processParagraphs(XWPFDocument document, Map<String, Object> paramData) {
        logger.info("process Paragraph");

        List<XWPFParagraph> paras = document.getParagraphs();

        if (paras == null) {
            return;
        }

        for (XWPFParagraph paragraph : paras) {
            doProcessParagraph(document, paragraph, paramData);
        }
    }

    private void doProcessParagraph(XWPFDocument document, XWPFParagraph paragraph, Map<String, Object> paramData) {

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
                logger.warn("No tag({}) handler found, skipped", tagName);
                return;
            }

            // 去掉 段落中的 tagName (tag 标识)
            TagParser.removeTagName(paragraph, tagName);


            // 格式化 碎片 run
            formatExpressionInMultiRuns(paragraph);


            // 调用handler对cell进行处理
            operateParagraphHandler.operate(document, paragraph, paramData, this.expressionCalculator);

        }

        /*String content = run.text();
        if (StringUtils.isNotEmpty(run.text())) {
            // 获取content中的标签
            String tagName = TagParser.findFirstTag(content);
            if (tagName == null) {
                return;
            }

            // 通过tag找到handler
            OperateParagraphHandler operateParagraphHandler = this.operateParagraphHandlerTagMap.get(tagName);
            if (operateParagraphHandler == null) {
                logger.warn("No tag({}) handler found, skipped", tagName);
                return;
            }

            // 调用handler对cell进行处理
            operateParagraphHandler.operate(document, run, paramData, this.expressionCalculator);
        }*/
    }

    /**
     * 校验并初始化参数列表（将一些公共的流程中的对象放入参数列表）
     *
     * @param paramData 参数
     */
    private void verifyAndInitParamData(Map<String, Object> paramData) {

        // 我们不允许用户使用这个key
        if (paramData.get(CommonDataNameConstant.ARR_HISTORY) != null) {
            logger.error("ARR_HISTORY cannot be used because some other information will be recorded");
            throw new IllegalArgumentException("ARR_HISTORY cannot be used because some other information will be recorded");
        }
        if (paramData.get(CommonDataNameConstant.MERGE_ARR_INFO) != null) {
            logger.error("MERGE_ARR_INFO cannot be used because some other information will be recorded");
            throw new IllegalArgumentException("MERGE_ARR_INFO cannot be used because some other information will be recorded");
        }
        // 初始化记录数组插入历史的全链路变量
        paramData.put(CommonDataNameConstant.ARR_HISTORY, new HashMap<String, ArrInfo>());
        paramData.put(CommonDataNameConstant.MERGE_ARR_INFO, new ArrayList<MergeArrInfo>());
    }

    /**
     * 构建器
     */
    public static class Builder {

        private final JxltEngine jxltEngine;

        private final Map<String, Object> functions;

        private int expressionCacheSize = 1000;

        private final List<OperateParagraphHandler> operateParagraphHandlerList;

        private final List<OperateTableCellHandler> operateTableCellHandlerList;

        private ExpressionCalculator expressionCalculator;


        /**
         * key: handler's tagName
         * value: OperateParagraphHandler
         */
        private Map<String, OperateParagraphHandler> operateParagraphHandlerTagMap;


        /**
         * key: handler's tagName
         * value: OperateTableCellHandler
         */
        private Map<String, OperateTableCellHandler> operateTableCellHandlerTagMap;


        private List<OperateParagraphHandlerHolder> operateParagraphHandlerHolderList;

        private List<OperateTableCellHandlerHolder> operateTableCellHandlerHolderList;

        public Builder() {
            functions = new HashMap<>();
            functions.putAll(loadFunctions());

            JexlEngine jexlEngine = new JexlBuilder()
                    .namespaces(functions)
                    .create();
            jxltEngine = jexlEngine.createJxltEngine();
            this.operateParagraphHandlerList = this.loadInstanceByInterface(OperateParagraphHandler.class);
            this.operateTableCellHandlerList = this.loadInstanceByInterface(OperateTableCellHandler.class);
            // 初始化辅助handler或者是map的信息
            this.initAuxiliaryInfo();

        }

        private void initAuxiliaryInfo() {
            this.initOperateParagraphHandlerHolderList();
            this.initOperateTableCellHandlerHolderList();
            this.initOperateParagraphHandlerTagMap();
            this.initOperateTableCellHandlerTagMap();
        }

        private void initOperateTableCellHandlerTagMap() {

            logger.info(" init OperateTableCellHandlerTagMap");
            this.operateTableCellHandlerTagMap = this.operateTableCellHandlerHolderList.stream()
                    .collect(Collectors.toConcurrentMap(
                            OperateTableCellHandlerHolder::getHandlerTagName,
                            OperateTableCellHandlerHolder::getOperateTableCellHandler,
                            (oldV, newV) -> {
                                throw new OperateWordHandlerInitException("There are multiple OperateHandler with the same tagName. cell:" + oldV + ";" + newV);
                            }
                    ));
        }

        private void initOperateTableCellHandlerHolderList() {
            logger.info("init OperateTableCellHandlerHolderList");
            this.operateTableCellHandlerHolderList = this.operateTableCellHandlerList.stream()
                    .map(cellHandler -> {
                        TagOperateHandler tagOperateHandler = cellHandler.getClass().getAnnotation(TagOperateHandler.class);
                        if (tagOperateHandler == null) {
                            throw new OperateWordHandlerInitException("The OperateHandler must identify the CellOperateHandler annotation");
                        }
                        return new OperateTableCellHandlerHolder(cellHandler, tagOperateHandler.tagName());

                    })
                    .collect(Collectors.toList());
        }

        private void initOperateParagraphHandlerHolderList() {
            logger.info(" init OperateParagraphHandlerHolderList");
            this.operateParagraphHandlerHolderList = this.operateParagraphHandlerList.stream()
                    .map(cellHandler -> {
                        TagOperateHandler tagOperateHandler = cellHandler.getClass().getAnnotation(TagOperateHandler.class);
                        if (tagOperateHandler == null) {
                            throw new OperateWordHandlerInitException("The OperateHandler must identify the CellOperateHandler annotation");
                        }
                        return new OperateParagraphHandlerHolder(cellHandler, tagOperateHandler.tagName());

                    })
                    .collect(Collectors.toList());
        }

        private void initOperateParagraphHandlerTagMap() {

            logger.info(" init OperateParagraphHandlerTagMap");
            this.operateParagraphHandlerTagMap = this.operateParagraphHandlerHolderList.stream()
                    .collect(Collectors.toConcurrentMap(
                            OperateParagraphHandlerHolder::getHandlerTagName,
                            OperateParagraphHandlerHolder::getOperateParagraphHandler,
                            (oldV, newV) -> {
                                throw new OperateWordHandlerInitException("There are multiple OperateHandler with the same tagName. cell:" + oldV + ";" + newV);
                            }
                    ));
        }


        private Map<String, Object> loadFunctions() {
            List<FunctionProvider> functionProviders = this.loadInstanceByInterface(FunctionProvider.class);
            return functionProviders.stream()
                    .map(FunctionProvider::provideFunctions)
                    .reduce(new HashMap<>(), (root, ele) -> {
                        root.putAll(ele);
                        return root;
                    });
        }


        private <T> List<T> loadInstanceByInterface(Class<T> clazz) {
            return SpiLoader.loadInstanceListSorted(clazz);
        }


        /**
         * 添加自定义函数
         *
         * @param key   自定义函数名称
         * @param clazz 自定义函数的实现类
         * @return builder
         */
        public Builder functions(String key, Class<?> clazz) {
            this.functions.put(key, clazz);
            return this;
        }


        /**
         * 设置表达式缓存map的容量
         *
         * @param cacheSize 大小
         */
        public Builder expressionCacheSize(int cacheSize) {
            this.expressionCacheSize = cacheSize;
            return this;
        }

        /**
         * 重新设置 表达式计算器
         */
        public Builder resetExpressionCalculator(ExpressionCalculator calculator) {
            this.expressionCalculator = calculator;
            return this;
        }

        /**
         * 构建
         */
        public DefaultTemplateWordFiller build() {
            return new DefaultTemplateWordFiller(this);
        }


    }
}
