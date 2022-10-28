package pers.lzy.template.word.pojo.poi;

import java.util.List;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/28  17:26
 */
public class TextWordCell {

    private final List<WordParagraph> paragraphs;


    public TextWordCell(List<WordParagraph> paragraphs) {
        this.paragraphs = paragraphs;
    }

    public List<WordParagraph> getParagraphs() {
        return paragraphs;
    }
}
