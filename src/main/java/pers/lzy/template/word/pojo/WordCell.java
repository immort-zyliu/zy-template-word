package pers.lzy.template.word.pojo;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import java.util.ArrayList;
import java.util.List;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  16:04
 */
public class WordCell {

    private final int cellIndex;

    private final XWPFTableCell cell;

    public WordCell(int cellIndex, XWPFTableCell cell) {
        this.cellIndex = cellIndex;
        this.cell = cell;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public XWPFTableCell getCell() {
        return cell;
    }

    public List<XWPFParagraph> paragraphs() {
        if (cell == null) {
            return new ArrayList<>();
        }

        return cell.getParagraphs();
    }
}
