package pers.lzy.template.word.provider;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/23  18:10
 */
@FunctionalInterface
public interface DocumentProvider {

    /**
     * sheet提供
     * @return 要操作的sheet
     */
    XWPFDocument getDocument();
}
