package pers.lzy.template.word.pojo;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;

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
}
