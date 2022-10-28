package pers.lzy.template.word.common;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import pers.lzy.template.word.utils.ReUtils;
import pers.lzy.template.word.utils.WordUtil;

import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static pers.lzy.template.word.utils.WordUtil.insertRunAndCopyStyle;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/24  11:06
 */
public class TagParser {

    private static final Logger log = LoggerFactory.getLogger(TagParser.class);

    /**
     * 匹配标签的正则
     * 标签可由 字母、数字、下划线、短横杠、空格组成
     */
    private static final Pattern TAG_PATTERN = Pattern.compile("(<|[</])([\\d\\w\\s-_]+?)>");

    private static final Pattern TAG_PATTERN_2 = Pattern.compile("^\\s*<(.+?)>:");

    private static final Pattern EXPRESSION_PATTERN = Pattern.compile("\\$\\{\\s*?(.+?)\\s*}");

    /**
     * 开始寻找表达式起始节点的标志
     */
    private static final int START_PHASE_FLAG = 0;

    /**
     * 正在寻找表达式结束（拼接表达式）的标志
     */
    private static final int FIND_PHASE_FLAG = 1;

    /**
     * 将tag name 还原成原始的tag
     *
     * @param tagName tagName
     * @return 结果
     */
    private static String restoreTag(String tagName) {

        return "<" + tagName + ">:";
    }

    /**
     * 校验里面是否含有 表达式 ${xxxx}
     *
     * @param content 内容
     * @return 结果true:含有，false：不含有
     */
    public static boolean verifyHasExpression(String content) {

        if (content == null) {
            return false;
        }

        Matcher matcher = EXPRESSION_PATTERN.matcher(content);

        return matcher.find();
    }


    /**
     * 从内容中获取第一个标签
     *
     * @param content 内容
     */
    public static String findFirstTag(String content) {
        List<String> tagList = ReUtils.findAll(TAG_PATTERN, content, 2);
        if (tagList.size() == 0) {
            return null;
        }
        return tagList.get(0);
    }

    public static String parseRunTagContent(XWPFRun run, String tagName) {
        if (run == null) {
            return null;
        }

        // 获取单元格中的所有内容
        String value = run.text();
        // 获取指定标签的内容
        String parsedValue = parseParamByTag(value, tagName);
        // 如果是空，说明不是 simple-eval标签，则不处理
        if (StringUtils.isEmpty(parsedValue)) {
            return null;
        }
        return parsedValue;
    }

    /**
     * 判断内容中是否含有指定的标签
     *
     * @return true：是，false：不是
     */
    public static String parseParamByTag(String content, String tagName) {
        // 构造正则
        String regStr = String.format("<%s>(.+)</%s>", tagName, tagName);
        // 只匹配第一个，因为我们只认第一个
        return ReUtils.get(regStr, content, 1);
    }

    /**
     * 获取内容的tag标签
     *
     * @param content 内容
     * @return 结果
     */
    public static String findContentTag(String content) {
        String tagName = ReUtils.get(TAG_PATTERN_2, content, 1);
        if (StringUtils.isBlank(tagName)) {
            return null;
        }
        return tagName;
    }

    /**
     * 整理段落，使得表达式在一个run中，避免出现表达式在不同的run中导致的表达式解析出错。
     *
     * @param paragraph 被整理的段落
     */
    public static void formatExpressionInMultiRuns(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null) {
            return;
        }


        // 0代表开始阶段，1代表寻找阶段
        int phase = START_PHASE_FLAG;

        // 记录前一个char
        char preChar = '-';

        // 记录最后一个 $ 符号所在 段落的 run  的 下标位置（会实时更新）
        int startCharRunIndex = 0;

        //记录run结束的角标，开始的角标为i
        int endExpressionRunIndex = 0;

        //包含占位符的字符缓存
        StringBuilder expressionCache = new StringBuilder();

        for (int curRunIndex = 0; curRunIndex < runs.size(); curRunIndex++) {
            XWPFRun outerCurRun = runs.get(curRunIndex);
            String outerCurRunText = outerCurRun.text();
            if (outerCurRunText == null) {
                continue;
            }

            for (int outerCurIndex = 0; outerCurIndex < outerCurRunText.length(); outerCurIndex++) {
                char outCurrentChar = outerCurRunText.charAt(outerCurIndex);

                if (outCurrentChar == '$') {
                    // 记录 表达式开始字符所在的 run的索引
                    startCharRunIndex = curRunIndex;
                }

                if (outCurrentChar == '{' && preChar == '$') {
                    // 找到了占位符的开始阶段

                    // 开始记录表达式
                    expressionCache = new StringBuilder();

                    // 如果 表达式开始字符所在的 run的索引 就是当前run的索引。说明 表达式开始字符在本run中.不用操作
                    // 否则需要将 表达式开始字符($) 所在的run也一并加入进来
                    if (startCharRunIndex != curRunIndex) {
                        expressionCache.append(runs.get(startCharRunIndex).text());
                    }

                    // 记录 此run 中 第二个开始字符({) 前面的字符缓存进表达式中
                    for (int forwardIndex = 0; forwardIndex <= outerCurIndex; forwardIndex++) {
                        expressionCache.append(outerCurRunText.charAt(forwardIndex));
                    }

                    // 记录寻找阶段已经开始
                    phase = FIND_PHASE_FLAG;
                    continue;
                }


                if (phase == FIND_PHASE_FLAG) {
                    // 如果现在所处的阶段是找寻阶段
                    // 则说明现在的字符都要记录在表达式中
                    expressionCache.append(outCurrentChar);

                    if (outCurrentChar == '}') {
                        // 说明表达式记录完毕

                        // 记录结束字符所在的run的索引
                        endExpressionRunIndex = curRunIndex;

                        // 阶段重新初始化为寻找阶段
                        phase = START_PHASE_FLAG;

                        // 将表达式，放入 $ 所在的run中
                        WordUtil.setRunValue(runs.get(startCharRunIndex), expressionCache.toString());


                        // 废弃run的起始位置( 当前索引位置是不删除的.)
                        int removeStartIndex = startCharRunIndex;

                        // 如果此run后续还有，
                        if (outerCurIndex != outerCurRunText.length() - 1) {

                            // 则复制出来一个run(格式复制此run的)，将剩余字符放入新run中
                            // 将run插入到 $所在run的后面。
                            insertRunAndCopyStyle(paragraph, outerCurIndex, startCharRunIndex + 1, outerCurRunText.substring(outerCurIndex + 1));
                            // 删除废弃的run的时候，不包含当前添加的run,所以遍历废弃run的时候，索引要注意一下
                            removeStartIndex++;
                        }


                        // 遍历删除废弃的run
                        for (int removeIndex = endExpressionRunIndex; removeIndex > removeStartIndex; removeIndex--) {
                            //角标移除后，runs会同步变动，直接继续处理i就可以
                            log.info("移除下标：{}", removeIndex);
                            paragraph.removeRun(removeIndex);
                        }

                    }

                }

                // 记录pre char
                preChar = outCurrentChar;
            }
        }

    }


    /**
     * 到这里，我们就已经默认您的段落满足tagName的要求，如果不满足，调用此方法会出错.
     *
     * @param paragraph 段落
     * @param tagName   要移除的 tagName
     */
    public static void removeTagName(XWPFParagraph paragraph, String tagName) {
        String tagFlag = restoreTag(tagName);
        int tagFlagLength = tagFlag.length();
        boolean breakFlag = false;
        boolean startFlag = false;

        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null) {
            return;
        }

        for (XWPFRun run : runs) {
            StringBuilder sb = new StringBuilder();
            String runText = run.text();
            if (runText != null) {
                for (int i = 0; i < runText.length(); i++) {
                    char curChar = runText.charAt(i);

                    // 如果是空格，还没有开启，则跳过
                    if ((Character.isSpaceChar(curChar) || Character.isWhitespace(curChar)) && !startFlag) {
                        sb.append(curChar);
                        continue;
                    }

                    // 如果不是空格，则需要开启
                    startFlag = true;

                    if (tagFlagLength == 0) {
                        // 说明需要将此run剩下的字符撞到sb中，然后重新设置进去
                        sb.append(curChar);
                        // 说明下一个run不用进行循环了，设置break为true
                        breakFlag = true;
                    }
                    tagFlagLength--;
                }
                if (tagFlagLength == 0) {
                    // 说明下一个run不用进行循环了，设置break为true
                    breakFlag = true;
                }

                // 遍历完成，将过滤后的字符放入run中
                run.setText(sb.toString(), 0);
            }

            if (breakFlag) {
                break;
            }
        }

    }

}
