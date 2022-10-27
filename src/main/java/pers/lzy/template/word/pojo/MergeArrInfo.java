package pers.lzy.template.word.pojo;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/28  13:10
 *
 * 记录时 arr-merge 标签的arr范围。
 */
public class MergeArrInfo {

    /**
     * 起始行
     */
    private Integer startRow;

    /**
     * 结束行
     */
    private Integer endRow;

    /**
     * 所在列
     */
    private Integer columnNumber;

    public MergeArrInfo() {
    }

    public MergeArrInfo(Integer startRow, Integer endRow, Integer columnNumber) {
        this.startRow = startRow;
        this.endRow = endRow;
        this.columnNumber = columnNumber;
    }

    public Integer getStartRow() {
        return startRow;
    }

    public void setStartRow(Integer startRow) {
        this.startRow = startRow;
    }

    public Integer getEndRow() {
        return endRow;
    }

    public void setEndRow(Integer endRow) {
        this.endRow = endRow;
    }

    public Integer getColumnNumber() {
        return columnNumber;
    }

    public void setColumnNumber(Integer columnNumber) {
        this.columnNumber = columnNumber;
    }
}
