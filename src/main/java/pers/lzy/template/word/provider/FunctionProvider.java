package pers.lzy.template.word.provider;

import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/7/11  17:34
 * 函数定义的提供者
 */
public interface FunctionProvider {

    /**
     * 提供function
     *
     * @return <pre>{                       </pre>
     * <pre>    key: functionName            </pre>
     * <pre>    value: 函数的类              </pre>
     * <pre>}                                </pre>
     */
    Map<String, Object> provideFunctions();
}
