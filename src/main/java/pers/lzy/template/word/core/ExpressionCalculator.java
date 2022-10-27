package pers.lzy.template.word.core;

import java.util.Collection;
import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/24  12:54
 * 表达式计算器
 */
public interface ExpressionCalculator {

    /**
     * 计算表达式。
     *
     * @param expression     表达式
     * @param calculateParam 计算所需要的参数
     * @return 计算后的结果
     */
    String calculate(String expression, Map<String, Object> calculateParam);

    /**
     * 计算表达式。
     *
     * @param expression     表达式
     * @param calculateParam 计算所需要的参数
     * @return 计算后的结果
     */
    Object calculateNoFormat(String expression, Map<String, Object> calculateParam);

    /**
     * 解析参数中的第一个数组，并返回此数组
     *
     * @param expressionStr  表达式字符串
     * @param calculateParam 计算参数
     * @return 解析出的数组
     */
    Collection<?> parseObjArrInParamMap(String expressionStr, Map<String, Object> calculateParam);

}
