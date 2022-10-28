package pers.lzy.template.word.test.function.provider;

import com.google.auto.service.AutoService;
import com.google.common.collect.Lists;
import pers.lzy.template.word.provider.FunctionProvider;
import pers.lzy.template.word.test.function.func.ScoreUtil;

import java.util.HashMap;
import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/7/11  17:52
 * 函数提供者
 */
@AutoService(FunctionProvider.class)
public class MyFunctionProvider3 implements FunctionProvider {


    /**
     * 提供function
     *
     * @return <pre>{                               </pre>
     * <pre>    key: functionName           </pre>
     * <pre>    value: 函数的类             </pre>
     * <pre>}                               </pre>
     */
    @Override
    public Map<String, Object> provideFunctions() {
        Map<String, Object> res = new HashMap<>();
        res.put("Lists", Lists.class);
        // 注册到全局
        res.put("ScoreUtil", ScoreUtil.class);
        return res;
    }
}
