package pers.lzy.template.word.pojo;

import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  16:03
 */
public class WordTable {

    private final int tableIndex;

    private final XWPFTable table;

    public WordTable(int tableIndex, XWPFTable table) {
        this.table = table;
        this.tableIndex = tableIndex;

    }

    public int getTableIndex() {
        return tableIndex;
    }

    public XWPFTable getTable() {
        return table;
    }
}
