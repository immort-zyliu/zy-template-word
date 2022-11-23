package pers.lzy.template.word.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pers.lzy.template.word.common.TagParser;
import pers.lzy.template.word.pojo.WordCell;
import pers.lzy.template.word.pojo.WordTable;
import pers.lzy.template.word.pojo.poi.TextWordCell;
import pers.lzy.template.word.pojo.poi.WordParagraph;
import pers.lzy.template.word.pojo.poi.WordRun;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/27  16:48
 */
public class WordUtil {

    private static final Logger log = LoggerFactory.getLogger(WordUtil.class);

    public static void setRunValue(XWPFRun run, Object value) {

        if (value == null) {
            run.setText(null, 0);
            return;
        }

        // 填充数据
        run.setText(value.toString(), 0);
    }

    /**
     * 清空段落 内容
     *
     * @param paragraph 段落
     */
    public static void cleanParagraphContent(XWPFParagraph paragraph) {
        for (XWPFRun run : paragraph.getRuns()) {
            run.setText(null, 0);
        }
    }

    /**
     * 插入一个run，并复制某个run的格式
     * <p>
     * 本方法不做任何校验，错了就正常抛出异常
     *
     * @param paragraph      被操作的段落
     * @param insertRunIndex 将这个run插入到 段落中的 run的索引
     * @param sourceRunIndex 要复制哪个run的 格式？
     * @param insertValue    被插入run中要存储哪些内容?
     */
    public static void insertRunAndCopyStyle(XWPFParagraph paragraph, int sourceRunIndex, int insertRunIndex, String insertValue) {

        List<XWPFRun> runs = paragraph.getRuns();
        XWPFRun sourceRun = runs.get(sourceRunIndex);

        XWPFRun insertRun = paragraph.insertNewRun(insertRunIndex);
        setRunValue(insertRun, insertValue);
        copyRunStyle(sourceRun, insertRun);


    }

    private static void copyRunStyle(XWPFRun sourceRun, XWPFRun targetRun) {
        // 复制格式

        STVerticalAlignRun.Enum verticalAlignment = sourceRun.getVerticalAlignment();
        if (verticalAlignment != null) {
            targetRun.setVerticalAlignment(verticalAlignment.toString());
        }

        targetRun.setBold(sourceRun.isBold());
        targetRun.setCapitalized(sourceRun.isCapitalized());
        targetRun.setCharacterSpacing(sourceRun.getCharacterSpacing());
        targetRun.setColor(sourceRun.getColor());
        targetRun.setDoubleStrikethrough(sourceRun.isDoubleStrikeThrough());
        targetRun.setEmbossed(sourceRun.isEmbossed());

        STEm.Enum emphasisMark = sourceRun.getEmphasisMark();
        if (emphasisMark != null) {
            targetRun.setEmphasisMark(emphasisMark.toString());
        }

        targetRun.setFontFamily(sourceRun.getFontFamily());
        if (sourceRun.getFontSize() != -1) {
            targetRun.setFontSize(sourceRun.getFontSize());
        }

        targetRun.setImprinted(sourceRun.isImprinted());
        targetRun.setItalic(sourceRun.isItalic());
        targetRun.setKerning(sourceRun.getKerning());
        targetRun.setLang(sourceRun.getLang());
        targetRun.setShadow(sourceRun.isShadowed());
        targetRun.setSmallCaps(sourceRun.isSmallCaps());
        targetRun.setStrikeThrough(sourceRun.isStrikeThrough());
        targetRun.setStyle(sourceRun.getStyle());
        // ####insertRun.setSubscript(sourceRun.getSubscript());

        STHighlightColor.Enum textHightlightColor = sourceRun.getTextHightlightColor();
        if (textHightlightColor != null) {
            targetRun.setTextHighlightColor(textHightlightColor.toString());
        }

        UnderlinePatterns underline = sourceRun.getUnderline();
        if (underline != null) {
            targetRun.setUnderline(underline);
        }


        targetRun.setVanish(sourceRun.isVanish());

        STThemeColor.Enum underlineThemeColor = sourceRun.getUnderlineThemeColor();
        if (underlineThemeColor != null) {
            targetRun.setUnderlineThemeColor(underlineThemeColor.toString());
        }


        targetRun.setUnderlineColor(sourceRun.getUnderlineColor());
        targetRun.setTextScale(sourceRun.getTextScale());
    }

    /**
     * 合并段落中的runs
     *
     * @param paragraph 段落
     */
    public static void mergeRunText(XWPFParagraph paragraph) {

        // 获取段落的合并字符串
        String text = paragraph.getText();
        List<XWPFRun> runs = paragraph.getRuns();
        // 将内容移动到第一个run中
        int index = 0;
        for (XWPFRun run : runs) {

            if (index == 0) {
                run.setText(null, 0);
                run.setText(text, 0);
            } else {
                run.setText(null, 0);
            }
            index++;
        }
    }

    /**
     * 写入
     */
    public static void writeFile(XWPFDocument xwpfDocument, File localModuleFile) {
        FileOutputStream excelFileOutPutStream = null;
        try {
            excelFileOutPutStream = new FileOutputStream(localModuleFile);
            xwpfDocument.write(excelFileOutPutStream);
            excelFileOutPutStream.flush();
        } catch (Exception e) {
            log.error("export error ", e);
        } finally {
            if (excelFileOutPutStream != null) {
                try {
                    excelFileOutPutStream.close();
                } catch (Exception e) {
                    log.error("export error ", e);
                }
            }
        }
    }

    /**
     * des:表末尾添加行(表，要复制样式的行，添加行数，插入的行下标索引)
     *
     * @param table          被操作下表格
     * @param sourceRowIndex 被复制的行
     * @param rows           要插入几行?
     * @param insertRowIndex 在那行后面进行插入?
     */
    public static void addRows(XWPFTable table, int sourceRowIndex, int rows, int insertRowIndex) {
        //循环添加行和和单元格
        for (int i = 1; i <= rows; i++) {
            //获取要复制样式的行
            XWPFTableRow sourceRow = table.getRow(sourceRowIndex);
            //添加新行
            XWPFTableRow targetRow = table.insertNewTableRow(insertRowIndex++);
            //复制行的样式给新行
            targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
            //获取要复制样式的行的单元格
            List<XWPFTableCell> sourceCells = sourceRow.getTableCells();
            //循环复制单元格
            for (XWPFTableCell sourceCell : sourceCells) {
                //添加新列
                XWPFTableCell newCell = targetRow.addNewTableCell();

                // 复制单元格的样式给新单元格
                newCell.setVerticalAlignment(sourceCell.getVerticalAlignment());
                newCell.setColor(sourceCell.getColor());
                newCell.setWidth(String.valueOf(sourceCell.getWidth()));
                newCell.setWidthType(sourceCell.getWidthType());
                newCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());

                //设置垂直居中
                // newCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                //复制单元格的居中方式给新单元格
                CTPPr pPr = sourceCell.getCTTc().getPList().get(0).getPPr();
                if (pPr != null && pPr.getJc() != null && pPr.getJc().getVal() != null) {
                    CTTc cttc = newCell.getCTTc();
                    CTP ctp = cttc.getPList().get(0);
                    CTPPr ctppr = ctp.getPPr();
                    if (ctppr == null) {
                        ctppr = ctp.addNewPPr();
                    }
                    CTJc ctjc = ctppr.getJc();
                    if (ctjc == null) {
                        ctjc = ctppr.addNewJc();
                    }
                    ctjc.setVal(pPr.getJc().getVal());
                }
                //得到复制单元格的段落
                    /*List<XWPFParagraph> sourceParagraphs = sourceCell.getParagraphs();
                    if (StringUtils.isEmpty(sourceCell.getText())) {
                        continue;
                    }
                    //拿到第一段
                    XWPFParagraph sourceParagraph = sourceParagraphs.get(0);
                    //得到新单元格的段落
                    List<XWPFParagraph> targetParagraphs = newCell.getParagraphs();
                    //判断新单元格是否为空
                    if (StringUtils.isEmpty(newCell.getText())) {
                        //添加新的段落
                        XWPFParagraph ph = newCell.addParagraph();
                        //复制段落样式给新段落
                        ph.getCTP().setPPr(sourceParagraph.getCTP().getPPr());
                        //得到文本对象
                        XWPFRun run = ph.getRuns().isEmpty() ? ph.createRun() : ph.getRuns().get(0);
                        //复制文本样式
                        run.setFontFamily(sourceParagraph.getRuns().get(0).getFontFamily());
                    } else {
                        XWPFParagraph ph = targetParagraphs.get(0);
                        ph.getCTP().setPPr(sourceParagraph.getCTP().getPPr());
                        XWPFRun run = ph.getRuns().isEmpty() ? ph.createRun() : ph.getRuns().get(0);
                        run.setFontFamily(sourceParagraph.getRuns().get(0).getFontFamily());
                    }*/
            }
        }
    }


    /**
     * @param table        表格
     * @param rowNum       目标行
     * @param colNum       目标列
     * @param textWordCell 处理好的 内容，直接替换即可
     * @param sourceCell   根据此单元格段落的格式 创建段落
     */
    public static void setCellObjValue(WordTable table, int rowNum, int colNum, TextWordCell textWordCell, WordCell sourceCell) {
        XWPFTable poiTable = table.getTable();
        XWPFTableRow row = poiTable.getRow(rowNum);
        if (row == null) {
            row = poiTable.insertNewTableRow(rowNum);
        }


        XWPFTableCell targetCell = row.getCell(colNum);

        List<XWPFParagraph> sourceParagraphList = sourceCell.getCell().getParagraphs();
        List<WordParagraph> textParagraphs = textWordCell.getParagraphs();

        for (int sourceParagraphIndex = 0; sourceParagraphIndex < sourceParagraphList.size(); sourceParagraphIndex++) {
            XWPFParagraph sourceParagraph = sourceParagraphList.get(sourceParagraphIndex);
            WordParagraph wordParagraph = textParagraphs.get(sourceParagraphIndex);


            XWPFParagraph targetParagraph = targetCell.getParagraphArray(sourceParagraphIndex);
            if (targetParagraph == null) {
                targetParagraph = targetCell.addParagraph();
            }

            // 复制段落格式....
            //复制段落样式给新段落
            targetParagraph.getCTP().setPPr(sourceParagraph.getCTP().getPPr());

            List<XWPFRun> sourceRuns = sourceParagraph.getRuns();
            if (sourceRuns == null) {
                continue;
            }


            List<WordRun> wordRuns = wordParagraph.getRuns();
            for (int sourceRunIndex = 0; sourceRunIndex < sourceRuns.size(); sourceRunIndex++) {

                // 复制格式用的
                XWPFRun sourceRun = sourceRuns.get(sourceRunIndex);
                // 复制内容用的
                WordRun wordRun = wordRuns.get(sourceRunIndex);

                XWPFRun targetRun = getRunByIndex(targetParagraph, sourceRunIndex);
                if (targetRun == null) {
                    targetRun = targetParagraph.createRun();
                }

                // 复制格式
                copyRunStyle(sourceRun, targetRun);

                // 复制值
                setRunValue(targetRun, wordRun.getValue());

            }
        }


    }

    private static XWPFRun getRunByIndex(XWPFParagraph targetParagraph, int index) {
        List<XWPFRun> runs = targetParagraph.getRuns();
        if (runs == null) {
            return null;
        }
        if (index >= 0 && index < runs.size()) {
            return runs.get(index);
        }
        return null;
    }

    /**
     * 抽取并 格式化 run 的cell
     *
     * @param cell 目标cell
     * @return 结果
     */
    public static TextWordCell extractFormatTextCell(WordCell cell, String tagName) {
        if (cell == null) {
            return null;
        }

        XWPFTableCell poiCell = cell.getCell();
        if (poiCell == null) {
            return null;
        }

        List<XWPFParagraph> paragraphs = poiCell.getParagraphs();
        if (paragraphs == null) {
            return null;
        }

        // 删除 tagName
        TagParser.removeTagName(paragraphs, tagName);

        List<WordParagraph> wordParagraphs = new ArrayList<>(paragraphs.size());

        List<WordRun> wordRuns;
        for (XWPFParagraph paragraph : paragraphs) {
            // 格式化paragraph
            TagParser.formatExpressionInMultiRuns(paragraph);

            List<XWPFRun> runs = paragraph.getRuns();
            if (runs == null) {
                wordParagraphs.add(new WordParagraph(null));
                continue;
            }

            wordRuns = new ArrayList<>();
            for (XWPFRun run : runs) {
                wordRuns.add(new WordRun(run.text()));
            }
            wordParagraphs.add(new WordParagraph(wordRuns));
        }

        return new TextWordCell(wordParagraphs);
    }


    /**
     * 跨列合并
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * 跨行合并
     * http://stackoverflow.com/questions/24907541/row-span-with-xwpftable
     */
    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        if (fromRow == toRow) {
            return;
        }
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * 获取单元格
     *
     * @param table    表格
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @return 值
     */
    public static String getTableCellValue(XWPFTable table, Integer rowIndex, Integer colIndex) {

        XWPFTableRow row = table.getRow(rowIndex);
        if (row == null) {
            return "";
        }

        XWPFTableCell cell = row.getCell(colIndex);
        if (cell == null) {
            return "";
        }
        return cell.getText();
    }
}
