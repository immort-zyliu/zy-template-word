package pers.lzy.template.word.processor;

import com.google.auto.service.AutoService;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pers.lzy.template.word.anno.HandlerOrder;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.core.OperateWordPostProcessor;
import pers.lzy.template.word.pojo.MergeArrInfo;
import pers.lzy.template.word.pojo.MergeRegion;
import pers.lzy.template.word.utils.WordUtil;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static pers.lzy.template.word.constant.CommonDataNameConstant.MERGE_ARR_INFO;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/11/22  17:53
 * 对word表格合并的后置处理
 */
@HandlerOrder(20000)
@AutoService(OperateWordPostProcessor.class)
public class OperateWordTableMergeRegionPostProcessor implements OperateWordPostProcessor {


    private static final Logger log = LoggerFactory.getLogger(OperateWordTableMergeRegionPostProcessor.class);


    /**
     * 对 填充完成之后 的 word 进行个性化处理的方法
     *
     * @param document             要操作的 document
     * @param params               需要的参数列表(当然，此数据可以在整个handler中流转)
     * @param expressionCalculator 表达式计算器
     */
    @Override
    public void operatePostProcess(XWPFDocument document, Map<String, Object> params, ExpressionCalculator expressionCalculator) {
        List<XWPFTable> tables = document.getTables();
        if (tables == null) {
            return;
        }

        for (int i = 0; i < tables.size(); i++) {
            XWPFTable table = tables.get(i);
            processTable(i, table, params);
        }
    }

    private void processTable(int tableIndex, XWPFTable table, Map<String, Object> params) {

        // 取出表格对应的数组信息。
        // key:tableIndex
        // value: MergeArrInfo
        @SuppressWarnings("unchecked")
        Map<Integer,List<MergeArrInfo>> mergeArrInfo = (Map<Integer, List<MergeArrInfo>>) params.get(MERGE_ARR_INFO);

        List<MergeArrInfo> mergeArrInfos = mergeArrInfo.get(tableIndex);

        // 没有此表格的合并信息，则跳过
        if (mergeArrInfos == null) {
            return;
        }

        for (MergeArrInfo arrInfo : mergeArrInfos) {
            // 扫描表格，开始合并
            doProcessTable(table, params, arrInfo);
        }

    }

    private void doProcessTable(XWPFTable table, Map<String, Object> params, MergeArrInfo arrInfo) {

        List<MergeRegion> mergeRegionList = new ArrayList<>();

        Integer startRow = arrInfo.getStartRow();
        Integer endRow = arrInfo.getEndRow();
        Integer columnNumber = arrInfo.getColumnNumber();


        // 首先默认第一个为 上一个单元格的值
        String preValue = WordUtil.getTableCellValue(table, startRow, columnNumber);
        // region默认是第一个单元格所在位置
        MergeRegion tempMergeRegion = new MergeRegion(
                startRow, startRow, columnNumber, columnNumber, preValue
        );

        for (int rowNumber = startRow; rowNumber < endRow; rowNumber++) {
            String currentValue = WordUtil.getTableCellValue(table, rowNumber, columnNumber);

            // 如果当前的 value 跟上一个不同,说明是新的开始
            if ("".equals(currentValue) && "".equals(preValue) || !currentValue.equals(preValue)) {
                // currentValue是空串且preValue是空串，那么我们就认为他不同（虽然他是相同的） true，则进来了
                // 如果到了|| 后面，说明前两个有一个不是空串，那么我们就有可比性了

                // 则结束封装并放入list，同时新来一个继续封装

                // 如果发现是多个单元格，才放入最终的集合中
                if (tempMergeRegion.getFirstRow() != tempMergeRegion.getLastRow()) {
                    mergeRegionList.add(tempMergeRegion);
                }

                // 创建一个新的，继续寻找
                tempMergeRegion = new MergeRegion(
                        rowNumber, rowNumber, columnNumber, columnNumber, currentValue
                );

                // 设置preValue 为 当前新的Value
                preValue = currentValue;
            } else {
                // 说明跟上个一样，则改变当前 region的范围
                tempMergeRegion.setLastRow(rowNumber);
            }

        }


        // 如果发现是多个单元格，才放入最终的集合中（最后一次合并的判断处理）
        if (tempMergeRegion.getFirstRow() != tempMergeRegion.getLastRow()) {
            mergeRegionList.add(tempMergeRegion);
        }

        // 进行真正的合并
        doMergeCellByMergeRegionInfo(table, mergeRegionList);

    }

    private void doMergeCellByMergeRegionInfo(XWPFTable table, List<MergeRegion> mergeRegionList) {
        // 遍历进行合并
        for (MergeRegion mergeRegion : mergeRegionList) {

            // 调用方法进行合并
            WordUtil.mergeCellsVertically(
                    table,
                    mergeRegion.getFirstCol(),
                    mergeRegion.getFirstRow(),
                    mergeRegion.getLastRow()
            );
        }
    }
}
