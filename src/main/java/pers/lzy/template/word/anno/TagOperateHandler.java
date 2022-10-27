package pers.lzy.template.word.anno;

import java.lang.annotation.*;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/23  17:15
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface TagOperateHandler {

    /**
     * The cell that identifies which tag this processor is used to process
     * -- zyliu
     * @return cell tag name
     */
    String tagName();
}
