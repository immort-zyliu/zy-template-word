package pers.lzy.template.word.pojo.poi;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  17:24
 */
public class WordRun {

    private String value;

    public WordRun(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }


    public void setValue(Object value) {
        if (value == null) {
            this.value = null;
            return;
        }
        this.value = value.toString();
    }
}
