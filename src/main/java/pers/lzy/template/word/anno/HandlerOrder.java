package pers.lzy.template.word.anno;

import java.lang.annotation.*;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/23  17:15
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface HandlerOrder {

    /**
     * Represents the lowest precedence.
     */
    int LOWEST_PRECEDENCE = Integer.MAX_VALUE;
    /**
     * Represents the highest precedence.
     */
    int HIGHEST_PRECEDENCE = Integer.MIN_VALUE;

    /**
     * The SPI precedence value. Lowest precedence by default.
     *
     * @return the precedence value
     */
    int value() default LOWEST_PRECEDENCE;
}
