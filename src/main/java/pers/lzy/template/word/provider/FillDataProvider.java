package pers.lzy.template.word.provider;

import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/24  9:53
 */
@FunctionalInterface
public interface FillDataProvider {

    /**
     * 提供数据，同时也做全局数据流转，
     */
    Map<String, Object> getParamData();
}
