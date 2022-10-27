package pers.lzy.template.word.common;


import java.util.regex.Pattern;

/**
 * 常用正则表达式集合(来自hutool工具类)
 */
public class PatternPool {

    /**
     * 占位符的正则：
     * eg :
     * ${name}
     *
     * author:zyliu-immort
     */
    public final static String RE_PLACEHOLDER = "\\$\\{(.+?)\\}";

    /**
     * 数组变量链的正则：
     * eg： school.people[].name
     * return: school.people
     *
     * author:zyliu-immort
     */
    public final static String RE_ARRAY_VARIABLE_CHAIN = "([a-zA-Z]\\w+\\.)*[a-zA-Z]\\w+\\[\\]";

    /**
     * 正则表达式匹配中文汉字
     */
    public final static String RE_CHINESE = "[\u4E00-\u9FFF]";
    /**
     * 正则表达式匹配中文字符串
     */
    public final static String RE_CHINESES = RE_CHINESE + "+";


    /**
     * 英文字母 、数字和下划线
     */
    public final static Pattern GENERAL = Pattern.compile("^\\w+$");
    /**
     * 数字
     */
    public final static Pattern NUMBERS = Pattern.compile("\\d+");
    /**
     * 字母
     */
    public final static Pattern WORD = Pattern.compile("[a-zA-Z]+");
    /**
     * 单个中文汉字
     */
    public final static Pattern CHINESE = Pattern.compile(RE_CHINESE);
    /**
     * 中文汉字
     */
    public final static Pattern CHINESES = Pattern.compile(RE_CHINESES);
    /**
     * 分组
     */
    public final static Pattern GROUP_VAR = Pattern.compile("\\$(\\d+)");

    // -------------------------------------------------------------------------------------------------------------------------------------------------------------------
    /**
     * Pattern池
     */
    private static final SimpleCache<RegexWithFlag, Pattern> POOL = new SimpleCache<>();

    /**
     * 先从Pattern池中查找正则对应的{@link Pattern}，找不到则编译正则表达式并入池。
     *
     * @param regex 正则表达式
     * @return {@link Pattern}
     */
    public static Pattern get(String regex) {
        return get(regex, 0);
    }

    /**
     * 先从Pattern池中查找正则对应的{@link Pattern}，找不到则编译正则表达式并入池。
     *
     * @param regex 正则表达式
     * @param flags 正则标识位集合 {@link Pattern}
     * @return {@link Pattern}
     */
    public static Pattern get(String regex, int flags) {
        final RegexWithFlag regexWithFlag = new RegexWithFlag(regex, flags);

        Pattern pattern = POOL.get(regexWithFlag);
        if (null == pattern) {
            pattern = Pattern.compile(regex, flags);
            POOL.put(regexWithFlag, pattern);
        }
        return pattern;
    }

    /**
     * 移除缓存
     *
     * @param regex 正则
     * @param flags 标识
     * @return 移除的{@link Pattern}，可能为{@code null}
     */
    public static Pattern remove(String regex, int flags) {
        return POOL.remove(new RegexWithFlag(regex, flags));
    }

    /**
     * 清空缓存池
     */
    public static void clear() {
        POOL.clear();
    }

    // ---------------------------------------------------------------------------------------------------------------------------------

    /**
     * 正则表达式和正则标识位的包装
     *
     * @author Looly
     */
    private static class RegexWithFlag {
        private final String regex;
        private final int flag;

        /**
         * 构造
         *
         * @param regex 正则
         * @param flag  标识
         */
        public RegexWithFlag(String regex, int flag) {
            this.regex = regex;
            this.flag = flag;
        }

        @Override
        public int hashCode() {
            final int prime = 31;
            int result = 1;
            result = prime * result + flag;
            result = prime * result + ((regex == null) ? 0 : regex.hashCode());
            return result;
        }

        @Override
        public boolean equals(Object obj) {
            if (this == obj) {
                return true;
            }
            if (obj == null) {
                return false;
            }
            if (getClass() != obj.getClass()) {
                return false;
            }
            RegexWithFlag other = (RegexWithFlag) obj;
            if (flag != other.flag) {
                return false;
            }
            if (regex == null) {
                return other.regex == null;
            } else {
                return regex.equals(other.regex);
            }
        }

    }
}
