package pers.lzy.template.word.utils;

import org.apache.commons.lang3.StringUtils;
import pers.lzy.template.excel.common.PatternPool;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 我直接复制的人家 hutool 里面的。
 *
 * @author hutool工具类
 * @since 2021/8/23
 */
public class ReUtils {


    /**
     * 取得内容中匹配的所有结果
     *
     * @param regex   正则
     * @param content 被查找的内容
     * @param group   正则的分组 0：整个匹配的内容，1 第一个括号里面的东西，后面依次类推
     * @return 结果列表
     * @since 3.0.6
     */
    public static List<String> findAll(String regex, CharSequence content, int group) {
        return findAll(regex, content, group, new ArrayList<>());
    }


    /**
     * 取得内容中匹配的所有结果
     *
     * @param <T>        集合类型
     * @param regex      正则
     * @param content    被查找的内容
     * @param group      正则的分组
     * @param collection 返回的集合类型
     * @return 结果集
     */
    public static <T extends Collection<String>> T findAll(String regex, CharSequence content, int group, T collection) {
        if (null == regex) {
            return collection;
        }

        return findAll(Pattern.compile(regex, Pattern.DOTALL), content, group, collection);
    }


    /**
     * 取得内容中匹配的所有结果
     *
     * @param <T>        集合类型
     * @param pattern    编译后的正则模式
     * @param content    被查找的内容
     * @param group      正则的分组
     * @param collection 返回的集合类型
     * @return 结果集
     */
    public static <T extends Collection<String>> T findAll(Pattern pattern, CharSequence content, int group, T collection) {
        if (null == pattern || null == content) {
            return null;
        }

        if (null == collection) {
            throw new NullPointerException("Null collection param provided!");
        }

        final Matcher matcher = pattern.matcher(content);
        while (matcher.find()) {
            collection.add(matcher.group(group));
        }
        return collection;
    }

    /**
     * 取得内容中匹配的所有结果
     *
     * @param pattern 编译后的正则模式
     * @param content 被查找的内容
     * @param group   正则的分组
     * @return 结果集
     */
    public static List<String> findAll(Pattern pattern, CharSequence content, int group) {
        if (null == pattern || null == content) {
            return null;
        }

        List<String> res = new ArrayList<>();

        final Matcher matcher = pattern.matcher(content);
        while (matcher.find()) {
            res.add(matcher.group(group));
        }
        return res;
    }


    /**
     * 获得匹配的字符串
     *
     * @param regex      匹配的正则
     * @param content    被匹配的内容
     * @param groupIndex 匹配正则的分组序号
     * @return 匹配后得到的字符串，未匹配返回null
     */
    public static String get(String regex, CharSequence content, int groupIndex) {
        if (null == content || null == regex) {
            return null;
        }

        // Pattern pattern = Pattern.compile(regex, Pattern.DOTALL);
        final Pattern pattern = PatternPool.get(regex, Pattern.DOTALL);
        return get(pattern, content, groupIndex);
    }


    /**
     * 获得匹配的字符串，对应分组0表示整个匹配内容，1表示第一个括号分组内容，依次类推
     *
     * @param pattern    编译后的正则模式
     * @param content    被匹配的内容
     * @param groupIndex 匹配正则的分组序号，0表示整个匹配内容，1表示第一个括号分组内容，依次类推
     * @return 匹配后得到的字符串，未匹配返回null
     */
    public static String get(Pattern pattern, CharSequence content, int groupIndex) {
        if (null == content || null == pattern) {
            return null;
        }

        final Matcher matcher = pattern.matcher(content);
        if (matcher.find()) {
            return matcher.group(groupIndex);
        }
        return null;
    }


    /**
     * 替换所有正则匹配的文本，并使用自定义函数决定如何替换
     *
     * @param str        要替换的字符串
     * @param regex      用于匹配的正则式
     * @param replaceFun 决定如何替换的函数
     * @return 替换后的文本
     */
    public static String replaceAll(CharSequence str, String regex, Func1<Matcher, String> replaceFun) {
        return replaceAll(str, Pattern.compile(regex), replaceFun);
    }


    public static String replaceAll(CharSequence str, Pattern pattern, Func1<Matcher, String> replaceFun) {
        if (StringUtils.isEmpty(str)) {
            return null == str ? null : str.toString();
        }

        final Matcher matcher = pattern.matcher(str);
        final StringBuffer buffer = new StringBuffer();
        while (matcher.find()) {
            try {
                matcher.appendReplacement(buffer, replaceFun.call(matcher));
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
        matcher.appendTail(buffer);
        return buffer.toString();
    }

    /**
     * 只有一个参数的函数对象<br>
     * 接口灵感来自于<a href="http://actframework.org/">ActFramework</a><br>
     * 一个函数接口代表一个一个函数，用于包装一个函数为对象<br>
     * 在JDK8之前，Java的函数并不能作为参数传递，也不能作为返回值存在，此接口用于将一个函数包装成为一个对象，从而传递对象
     *
     * @param <P> 参数类型
     * @param <R> 返回值类型
     * @author Looly
     * @since 4.2.2
     */
    @FunctionalInterface
    public interface Func1<P, R> {

        /**
         * 执行函数
         *
         * @param parameter 参数
         * @return 函数执行结果
         * @throws Exception 自定义异常
         */
        R call(P parameter) throws Exception;
    }
}
