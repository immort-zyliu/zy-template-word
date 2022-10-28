package pers.lzy.template.word.test.function.provider;

import com.google.auto.service.AutoService;
import org.apache.commons.lang3.StringUtils;
import pers.lzy.template.word.provider.FunctionProvider;

import java.util.HashMap;
import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/7/11  17:47
 */
@AutoService(FunctionProvider.class)
public class MyFunctionProvider2 implements FunctionProvider {


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
        res.put("StringUtils", StringUtils.class);
        return res;
    }
}
