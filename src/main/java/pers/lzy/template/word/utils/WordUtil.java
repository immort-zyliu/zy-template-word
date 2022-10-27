package pers.lzy.template.word.utils;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
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
        try {
            //获取表格的总行数
            int index = table.getNumberOfRows();
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
                    //复制单元格的样式给新单元格
                    newCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
                    //设置垂直居中
                    newCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);//垂直居中
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
                        ctjc.setVal(pPr.getJc().getVal()); //水平居中
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
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage());
        }
    }
}
