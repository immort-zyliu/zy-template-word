package pers.lzy.template.word.core;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pers.lzy.template.word.provider.DocumentProvider;
import pers.lzy.template.word.provider.FillDataProvider;

import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/24  10:03
 */
public interface TemplateWordFiller {

    /**
     * 给sheet填充对应的数据，根据模板
     *
     * @param documentProvider doc 提供者
     * @param dataProvider     数据提供着
     */
    void fillData(DocumentProvider documentProvider, FillDataProvider dataProvider);


    /**
     * 给sheet填充对应的数据，根据模板
     *
     * @param document  被填充的  doc
     * @param paramData 数据
     */
    default void fillData(XWPFDocument document, Map<String, Object> paramData) {
        fillData(() -> document, () -> paramData);
    }
}
