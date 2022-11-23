package pers.lzy.template.word.pojo;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/28  13:10
 * <p>
 * 存储了需要合并的单元格的范围信息。
 */
public class MergeRegion {
    private int firstRow;
    private int lastRow;
    private int firstCol;
    private int lastCol;
    private String val;


    public MergeRegion() {
    }

    public MergeRegion(int firstRow, int lastRow, int firstCol, int lastCol, String val) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;
        this.val = val;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public void setFirstRow(int firstRow) {
        this.firstRow = firstRow;
    }

    public int getLastRow() {
        return lastRow;
    }

    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    public int getFirstCol() {
        return firstCol;
    }

    public void setFirstCol(int firstCol) {
        this.firstCol = firstCol;
    }

    public int getLastCol() {
        return lastCol;
    }

    public void setLastCol(int lastCol) {
        this.lastCol = lastCol;
    }

    public String getVal() {
        return val;
    }

    public void setVal(String val) {
        this.val = val;
    }
}