package pers.lzy.template.word.pojo;

import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  16:04
 */
public class WordRow {

    private final int rowIndex;

    private final XWPFTableRow row;


    public WordRow(int rowIndex, XWPFTableRow row) {
        this.rowIndex = rowIndex;
        this.row = row;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public XWPFTableRow getRow() {
        return row;
    }
}
