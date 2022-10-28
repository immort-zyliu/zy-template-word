package pers.lzy.template.word.pojo.poi;

import java.util.List;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  17:25
 */
public class WordParagraph {

    private final List<WordRun> runs;

    public WordParagraph(List<WordRun> runs) {
        this.runs = runs;
    }

    public List<WordRun> getRuns() {
        return runs;
    }
}
