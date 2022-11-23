package pers.lzy.template.word.calculator;

import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.jexl3.JxltEngine;
import org.apache.commons.jexl3.MapContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pers.lzy.template.word.common.PatternPool;
import pers.lzy.template.word.exception.CalculateException;
import pers.lzy.template.word.core.ExpressionCalculator;
import pers.lzy.template.word.utils.ReUtils;

import javax.annotation.Nonnull;
import java.util.*;
import java.util.concurrent.ExecutionException;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/24  12:57
 * 基于jxel3的表达式计算实现
 */
public class Jxel3ExpressionCalculator implements ExpressionCalculator {

    private final static Logger logger = LoggerFactory.getLogger(Jxel3ExpressionCalculator.class);


    private final JxltEngine jxltEngine;


    /**
     * JxltEngine.Expression 是否是线程安全的呢？
     * 从 其 的 javadoc 中并没有发现 “线程不安全”的字眼
     * 并且，经我测试，其也是线程安全的，所以说可以给 expression 对象做公共缓存
     * by zyliu
     */
    private final LoadingCache<String, JxltEngine.Expression> READ_WRITE_CACHE_MAP;


    public Jxel3ExpressionCalculator(JxltEngine jxltEngine, int expressionCacheSize) {
        this.jxltEngine = jxltEngine;
        int guavaCacheSize;
        // 我们最低就是100，如果你传进来的比一百还小，那么默认就是1000
        if (expressionCacheSize >= 100) {
            guavaCacheSize = expressionCacheSize;
        } else {
            guavaCacheSize = 1000;
        }


        // 初始化表达式的cacheMap
        READ_WRITE_CACHE_MAP = CacheBuilder.newBuilder()
                .maximumSize(guavaCacheSize)
                .build(new CacheLoader<String, JxltEngine.Expression>() {
                    @Override
                    public JxltEngine.Expression load(@Nonnull String expressionStr) {
                        return Jxel3ExpressionCalculator.this.jxltEngine.createExpression(expressionStr);
                    }
                });
    }

    /**
     * 根据表达式字符串获取对应的 表达式对象
     *
     * @param expressionStr 表达式 字符串
     * @return 表达式
     */
    public JxltEngine.Expression getJxltEngineExpression(String expressionStr) throws ExecutionException {
        return READ_WRITE_CACHE_MAP.get(expressionStr);
    }


    /**
     * 计算表达式。
     *
     * @param expressionStr  表达式
     * @param calculateParam 计算所需要的参数
     * @return 计算后的结果
     */
    @Override
    public String calculate(String expressionStr, Map<String, Object> calculateParam) {
        // 获取参数
        // 进行表达式计算。
        String result = "";
        try {
            // 获取表达式，计算返回。
            JxltEngine.Expression expression = READ_WRITE_CACHE_MAP.get(expressionStr);
            Object evaluateRes = expression.evaluate(new MapContext(calculateParam));
            if (evaluateRes != null) {
                result = evaluateRes.toString();
            }
        } catch (ExecutionException e) {
            logger.error("Description Failed to get an expression from cache:", e);
            throw new CalculateException("Description Failed to get an expression from cache");
        }
        return result;
    }

    /**
     * 计算表达式。
     *
     * @param expressionStr  表达式
     * @param calculateParam 计算所需要的参数
     * @return 计算后的结果
     */
    @Override
    public Object calculateNoFormat(String expressionStr, Map<String, Object> calculateParam) {
        // 获取参数
        // 进行表达式计算。
        Object evaluateRes;
        try {
            // 获取表达式，计算返回。
            JxltEngine.Expression expression = READ_WRITE_CACHE_MAP.get(expressionStr);
            evaluateRes = expression.evaluate(new MapContext(calculateParam));
        } catch (ExecutionException e) {
            logger.error("Description Failed to get an expression from cache:", e);
            throw new CalculateException("Description Failed to get an expression from cache");
        }
        return evaluateRes;
    }

    /**
     * 返回参数中最短的那个数组
     *
     * @param expressionStr  表达式字符串
     * @param calculateParam 计算参数
     * @return 解析出的数组
     */
    @Override
    public Collection<?> parseObjArrInParamMap(String expressionStr, Map<String, Object> calculateParam) {
        Set<String> objExprSet = this.getObjExpr(expressionStr);

        if (CollectionUtils.isEmpty(objExprSet)) {
            logger.warn("Array expression not found...eg : school.people[]");
            return new ArrayList<>();
        }

        int minSize = Integer.MAX_VALUE;
        Collection<?> minSizeCollection = new ArrayList<>();

        for (String objExpr : objExprSet) {
            Object property = jxltEngine.getEngine().getProperty(calculateParam, objExpr);
            if (!(property instanceof Collection)) {
                logger.warn("Expression maps are not subclasses of Collection");
                continue;
            }
            int currentSize = ((Collection<?>) property).size();
            if (currentSize < minSize) {
                // 记录最小的property 和 minSize
                minSizeCollection = (Collection<?>) property;
                minSize = currentSize;
            }
        }

        return minSizeCollection;
    }

    /**
     * 解析表达式，并返回所有的数组变量的标识
     * eg： school.class[].name
     * 则会返回 school.class
     *
     * @param expressionStr 字符串表达式
     * @return 数组表达式
     */
    private Set<String> getObjExpr(String expressionStr) {
        Set<String> objExprSet = new HashSet<>();
        // 获取表达式中的内容
        List<String> foundAllPath = ReUtils.findAll(PatternPool.RE_PLACEHOLDER, expressionStr, 1);

        // 解析表达式中的内容。
        for (String allPath : foundAllPath) {
            List<String> allMatch = ReUtils.findAll(PatternPool.RE_ARRAY_CHINESE_VARIABLE_CHAIN, allPath, 0);
            // 添加的同时，去掉"[]"
            allMatch.forEach(ele -> objExprSet.add(ele.replaceAll("\\[]", "")));
        }
        return objExprSet;
    }


}
