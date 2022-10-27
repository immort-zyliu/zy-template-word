package pers.lzy.template.word.pojo;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/25  11:12
 * <p>
 * 记录数组单元格插入信息的
 */
public class ArrInfo {

    private Integer startRow;

    /**
     * 数组长度
     */
    private Integer size = 0;

    /**
     * 是否需要插入行的标志
     * :true 需要插入
     */
    private Boolean insertRowFlag = true;


    /**
     * 最小列
     */
    private Integer minColumnIndex = 0;

    /**
     * 设置最小的minColumnIndex;
     *
     * @param columnIndex 列索引
     */
    public void setMinColumnIndexFromGiven(int columnIndex) {
        if (columnIndex < this.minColumnIndex) {
            this.minColumnIndex = columnIndex;
        }
    }

    public Integer getStartRow() {
        return startRow;
    }

    public void setStartRow(Integer startRow) {
        this.startRow = startRow;
    }

    public Integer getSize() {
        return size;
    }

    public void setSize(Integer size) {
        this.size = size;
    }

    public Boolean getInsertRowFlag() {
        return insertRowFlag;
    }

    public void setInsertRowFlag(Boolean insertRowFlag) {
        this.insertRowFlag = insertRowFlag;
    }

    public Integer getMinColumnIndex() {
        return minColumnIndex;
    }

    public void setMinColumnIndex(Integer minColumnIndex) {
        this.minColumnIndex = minColumnIndex;
    }
}
