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
import pers.lzy.template.word.core.DocumentParagraphFiller;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.core.OperateWordPostProcessor;
import pers.lzy.template.word.core.TemplateWordFiller;
import pers.lzy.template.word.core.filler.DefaultDocumentParagraphFiller;
import pers.lzy.template.word.core.handler.OperateParagraphHandler;
import pers.lzy.template.word.core.handler.OperateTableCellHandler;
import pers.lzy.template.word.core.holder.OperateParagraphHandlerHolder;
import pers.lzy.template.word.core.holder.OperateTableCellHandlerHolder;
import pers.lzy.template.word.exception.OperateWordHandlerInitException;
import pers.lzy.template.word.pojo.*;
import pers.lzy.template.word.provider.DocumentProvider;
import pers.lzy.template.word.provider.FillDataProvider;
import pers.lzy.template.word.provider.FunctionProvider;
import pers.lzy.template.word.utils.SpiLoader;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static pers.lzy.template.word.constant.CommonDataNameConstant.ARR_HISTORY;
import static pers.lzy.template.word.constant.CommonDataNameConstant.ARR_HISTORY_ITEM;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/27  16:54
 */
public class DefaultTemplateWordFiller implements TemplateWordFiller {

    private static final Logger logger = LoggerFactory.getLogger(DefaultTemplateWordFiller.class);

    /**
     * 文档段落填充器
     */
    private final DocumentParagraphFiller documentParagraphFiller;

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

    private final List<OperateWordPostProcessor> operateWordPostProcessorList;

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
        this.documentParagraphFiller = DefaultDocumentParagraphFiller.getInstance(operateParagraphHandlerTagMap, this.expressionCalculator);
        this.operateWordPostProcessorList = builder.operateWordPostProcessorList;
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

        postProcessWord(document, paramData);
    }

    private void postProcessWord(XWPFDocument document, Map<String, Object> paramData) {
        logger.info("post process word doc");
        operateWordPostProcessorList
                .forEach(postProcessor -> postProcessor.operatePostProcess(document, paramData, this.expressionCalculator));
    }

    private void processTables(XWPFDocument document, Map<String, Object> paramData) {
        logger.info("process table");
        List<XWPFTable> tables = document.getTables();//获取表格

        if (tables == null) {
            return;
        }

        WordTable wordTable;
        WordRow wordRow;
        WordCell wordCell;

        for (int tableIndex = 0; tableIndex < tables.size(); tableIndex++) {
            XWPFTable table = tables.get(tableIndex);
            wordTable = new WordTable(tableIndex, table);

            List<XWPFTableRow> rows = table.getRows();

            // 初始化当前表格 所用参数
            initCurrentTableParam(paramData, tableIndex);


            if (rows == null) {
                continue;
            }

            for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
                XWPFTableRow row = rows.get(rowIndex);
                wordRow = new WordRow(rowIndex, row);

                List<XWPFTableCell> tableCells = row.getTableCells();
                if (tableCells == null) {
                    return;
                }

                for (int cellIndex = 0; cellIndex < tableCells.size(); cellIndex++) {
                    XWPFTableCell cell = tableCells.get(cellIndex);
                    wordCell = new WordCell(cellIndex, cell);

                    if (cell != null) {
                        doProcessCell(document, wordTable, wordRow, wordCell, paramData);
                    }
                }

            }

        }

    }

    /**
     * 初始化当前表格所用参数
     *
     * @param paramData  全局数据
     * @param tableIndex 当前表格索引
     */
    private void initCurrentTableParam(Map<String, Object> paramData, int tableIndex) {
        Map<String, ArrInfo> tableArrItemInfo = new HashMap<>();
        @SuppressWarnings("unchecked")
        Map<Integer, Map<String, ArrInfo>> arrAyHistory = (Map<Integer, Map<String, ArrInfo>>) paramData.get(ARR_HISTORY);
        arrAyHistory.put(tableIndex, tableArrItemInfo);
        // 设置当前表格的 使用的 arr的history
        paramData.put(ARR_HISTORY_ITEM, tableArrItemInfo);
    }

    private void doProcessCell(XWPFDocument document, WordTable table, WordRow row, WordCell cell, Map<String, Object> paramData) {

        // 获取cell中的内容
        String content = cell.getCell().getText();
        if (StringUtils.isNotBlank(content)) {
            // 获取content中的标签
            String tagName = TagParser.findContentTag(content);
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
            operateTableCellHandler.operate(document, table, row, cell, paramData, this.expressionCalculator, this.documentParagraphFiller);
        }

    }

    private void processParagraphs(XWPFDocument document, Map<String, Object> paramData) {
        logger.info("process Paragraph");

        List<XWPFParagraph> paras = document.getParagraphs();

        if (paras == null) {
            return;
        }

        paras.forEach(paragraph -> doProcessParagraph(document, paragraph, paramData));
    }

    private void doProcessParagraph(XWPFDocument document, XWPFParagraph paragraph, Map<String, Object> paramData) {
        documentParagraphFiller.doProcessParagraph(document, paragraph, paramData);
    }

    /**
     * 校验并初始化参数列表（将一些公共的流程中的对象放入参数列表）
     *
     * @param paramData 参数
     */
    private void verifyAndInitParamData(Map<String, Object> paramData) {

        // 我们不允许用户使用这个key
        if (paramData.get(ARR_HISTORY) != null) {
            logger.error("ARR_HISTORY cannot be used because some other information will be recorded");
            throw new IllegalArgumentException("ARR_HISTORY cannot be used because some other information will be recorded");
        }
        if (paramData.get(CommonDataNameConstant.MERGE_ARR_INFO) != null) {
            logger.error("MERGE_ARR_INFO cannot be used because some other information will be recorded");
            throw new IllegalArgumentException("MERGE_ARR_INFO cannot be used because some other information will be recorded");
        }

        if (paramData.get(ARR_HISTORY_ITEM) != null) {
            logger.error("ARR_HISTORY_ITEM cannot be used because some other information will be recorded");
            throw new IllegalArgumentException("ARR_HISTORY_ITEM cannot be used because some other information will be recorded");
        }
        // 初始化记录数组插入历史的全链路变量
        // key：tableIndex，vlale:{ key: rowIndex, value:ArrayInfo}
        paramData.put(ARR_HISTORY, new HashMap<Integer, Map<String, ArrInfo>>());
        paramData.put(ARR_HISTORY_ITEM, new HashMap<String, ArrInfo>());
        // key:tableIndex
        // value: MergeArrInfo
        paramData.put(CommonDataNameConstant.MERGE_ARR_INFO, new HashMap<Integer,List<MergeArrInfo>>());
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

        private final List<OperateWordPostProcessor> operateWordPostProcessorList;

        public Builder() {
            functions = new HashMap<>();
            functions.putAll(loadFunctions());

            JexlEngine jexlEngine = new JexlBuilder()
                    .namespaces(functions)
                    .create();
            jxltEngine = jexlEngine.createJxltEngine();
            this.operateParagraphHandlerList = this.loadInstanceByInterface(OperateParagraphHandler.class);
            this.operateTableCellHandlerList = this.loadInstanceByInterface(OperateTableCellHandler.class);
            this.operateWordPostProcessorList = this.loadInstanceByInterface(OperateWordPostProcessor.class);
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

            logger.info("init OperateTableCellHandlerTagMap");
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
            logger.info("init OperateParagraphHandlerHolderList");
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

            logger.info("init OperateParagraphHandlerTagMap");
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
