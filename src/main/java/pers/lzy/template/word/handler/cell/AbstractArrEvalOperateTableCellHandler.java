package pers.lzy.template.word.handler.cell;

import com.alibaba.fastjson.JSON;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pers.lzy.template.word.constant.CommonDataNameConstant;
import pers.lzy.template.word.core.AbstractOperateTableCellHandler;
import pers.lzy.template.word.core.DocumentParagraphFiller;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.pojo.*;
import pers.lzy.template.word.pojo.poi.TextWordCell;
import pers.lzy.template.word.pojo.poi.WordParagraph;
import pers.lzy.template.word.pojo.poi.WordRun;
import pers.lzy.template.word.utils.WordUtil;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import static pers.lzy.template.word.common.TagParser.verifyHasExpression;
import static pers.lzy.template.word.constant.CommonDataNameConstant.MERGE_ARR_INFO;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  13:36
 */
public abstract class AbstractArrEvalOperateTableCellHandler extends AbstractOperateTableCellHandler {


    /**
     * 对 word 表格中的 cell 进行个性化处理的方法
     *
     * @param document                被操作的文档
     * @param table                   被操作的表格
     * @param row                     被操作的行
     * @param cell                    被操作的单元格
     * @param params                  需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionCalculator    表达式计算器
     * @param documentParagraphFiller 文档段落填充器
     */
    @Override
    public void operate(XWPFDocument document, WordTable table, WordRow row, WordCell cell, Map<String, Object> params, ExpressionCalculator expressionCalculator, DocumentParagraphFiller documentParagraphFiller) {
        String cellText = cell.getCell().getText();

        // 获取目标数组
        Collection<?> arrInParamsByExpression = expressionCalculator.parseObjArrInParamMap(cellText, params);

        // 获取数组标记量
        @SuppressWarnings("unchecked")
        Map<String, ArrInfo> arrHistory = (Map<String, ArrInfo>) params.get(CommonDataNameConstant.ARR_HISTORY_ITEM);

        // 初始化遍历容量
        int traverseNumber = arrInParamsByExpression.size();

        // 根据数组的长度插入行(如果需要的话)
        int rowIndex = row.getRowIndex();

        // 获取arrInfo
        ArrInfo arrInfo = parseArrInfo(arrHistory, rowIndex, cell, arrInParamsByExpression);

        // 如果是arr-merge 标签的话，需要将当前的 arrInfo 放入,以便后面合并的时候使用。
        saveMargeArrInfoIfNecessary(table, arrInfo, cell, params);


        // 确定需要遍历的行数
        traverseNumber = determineTraverseNumber(traverseNumber, arrInfo);

        // 插入行，如果需要的话
        insertRowsIfNecessary(table, cell, arrInfo);

        // 抽取cell,并格式化run
        TextWordCell textWordCell = WordUtil.extractFormatTextCell(cell, this.tagName);

        // 对这一列进行数据的填充(吧段落抽出来)
        fillArrDataToTable(table, cell, textWordCell, traverseNumber, arrInfo, params, expressionCalculator);

    }

    private void saveMargeArrInfoIfNecessary(WordTable table, ArrInfo arrInfo, WordCell cell, Map<String, Object> params) {
        // 决定是否需要合并arr 单元格
        if (!mergeArrCell()) {
            return;
        }

        // key:tableIndex
        // value: MergeArrInfo
        @SuppressWarnings("unchecked")
        Map<Integer,List<MergeArrInfo>> mergeArrInfo = (Map<Integer, List<MergeArrInfo>>) params.get(MERGE_ARR_INFO);

        List<MergeArrInfo> tableMergeArrInfoList = mergeArrInfo.computeIfAbsent(table.getTableIndex(), k -> new ArrayList<>());

        tableMergeArrInfoList.add(
                new MergeArrInfo(
                        arrInfo.getStartRow(),
                        arrInfo.getStartRow() + arrInfo.getSize(),
                        cell.getCellIndex()
                )
        );

    }

    /**
     * 是否需要合并相同单元格
     *
     * @return true:合并，false不合并
     */
    protected abstract boolean mergeArrCell();


    private void fillArrDataToTable(WordTable table, WordCell cell, TextWordCell textWordCell, int traverseNumber, ArrInfo arrInfo, Map<String, Object> params, ExpressionCalculator expressionCalculator) {

        boolean blankFlag = false;

        if (traverseNumber <= 0) {
            // 此时就说明没有数组数据，或者是数组元素个数为0
            // 我们就要设置 空 flag 为true; 同时让其遍历一次，使其将当前单元格制空（因为没有值嘛）
            traverseNumber ++;
            blankFlag = true;
        }

        for (int index = 0; index < traverseNumber; index++) {

            // 克隆textWordCell，并使用表达式计算内容
            TextWordCell processedTextWordCell = this.cloneAndCalculate(textWordCell, expressionCalculator, params, index, blankFlag);


            // 将克隆，计算后的 textWordCell，弄到新的单元格中。
            // 单元格段落的格式，遵循模板cell中的段落。
            WordUtil.setCellObjValue(table,
                    arrInfo.getStartRow() + index,
                    cell.getCellIndex(),
                    processedTextWordCell,
                    cell
            );
        }
    }

    private TextWordCell cloneAndCalculate(TextWordCell textWordCell, ExpressionCalculator expressionCalculator, Map<String, Object> params, int paramArrIndex, boolean blankFlag) {

        TextWordCell textWordCellClone = JSON.parseObject(JSON.toJSONString(textWordCell), TextWordCell.class);

        List<WordParagraph> paragraphs = textWordCellClone.getParagraphs();
        if (paragraphs == null) {
            return textWordCellClone;
        }

        for (WordParagraph paragraph : paragraphs) {
            List<WordRun> runs = paragraph.getRuns();
            if (runs == null) {
                continue;
            }

            for (WordRun run : runs) {
                if (blankFlag) {
                    // 说明要填充为空
                    run.setValue(this.formatCellValue(null));
                    // 本次循环结束，让其进行下一次循环
                    continue;
                }

                String expression = run.getValue();
                // 说明没有解析出来表达式
                if (!verifyHasExpression(expression)) {
                    continue;
                }
                // 将expressionStr中的[] 中填充数字，方便取出数据
                String realExpressionStr = expression.replaceAll("\\[]", String.format("[%d]", paramArrIndex));

                // 进行表达式计算
                // 说明需要处理, 计算表达式并赋值。
                Object result = expressionCalculator.calculateNoFormat(realExpressionStr, params);
                // 设置到单元格中
                run.setValue(this.formatCellValue(result));
            }
        }


        return textWordCellClone;
    }

    private Object formatCellValue(Object realValue) {
        return realValue;
    }

    private void insertRowsIfNecessary(WordTable table, WordCell cell, ArrInfo arrInfo) {
        // 第一次是需要插入行的。
        // 3. 判断是否应该插入行，如果成立，则进行行的插入。
        if (arrInfo.getInsertRowFlag()) {
            // 进行行的插入
            WordUtil.addRows(table.getTable(), arrInfo.getStartRow(), arrInfo.getSize() - 1, arrInfo.getStartRow() + 1);
            //合并左侧的单元格，如资金详细信息左侧有个合并的资金信息的标题
            // 当然啦，只有左侧是合并单元格的时候，才会合并左侧的单元格
            // 还有一个条件，就是这是再插入arrMerge标签的时候
            // ExcelUtil.mergeLeft(sheet, cell, arrInfo.getSize() - 1);
        }
    }

    private int determineTraverseNumber(int traverseNumber, ArrInfo arrInfo) {
        // 确定第一次
        Integer arrTraverseNumber = arrInfo.getSize();
        // 确定 最小的遍历容量。
        return Math.min(traverseNumber, arrTraverseNumber);
    }

    private ArrInfo parseArrInfo(Map<String, ArrInfo> arrHistory, int rowIndex, WordCell cell, Collection<?> arrInParamsByExpression) {
        ArrInfo arrInfo = arrHistory.get(String.valueOf(rowIndex));
        if (arrInfo == null) {
            // 说明需要插入行
            arrInfo = new ArrInfo();
            arrInfo.setStartRow(rowIndex);
            arrInfo.setSize(arrInParamsByExpression.size());
            arrInfo.setInsertRowFlag(true);
            arrInfo.setMinColumnIndexFromGiven(cell.getCellIndex());
            arrHistory.put(String.valueOf(rowIndex), arrInfo);
        } else {
            // 说明是其他列的，则试探一下最小（不过一般是不会成功的,因为第一次设置的那个就是最小的）
            arrInfo.setMinColumnIndexFromGiven(cell.getCellIndex());
            // 说明已经插入过行了。
            arrInfo.setInsertRowFlag(false);
        }
        return arrInfo;
    }
}
