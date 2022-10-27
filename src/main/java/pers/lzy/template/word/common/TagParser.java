package pers.lzy.template.word.common;

import pers.lzy.template.word.utils.ReUtils;

import java.util.List;
import java.util.regex.Pattern;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/24  11:06
 */
public class TagParser {

    /**
     * 匹配标签的正则
     * 标签可由 字母、数字、下划线、短横杠、空格组成
     */
    private static final Pattern TAG_PATTERN = Pattern.compile("(<|[</])([\\d\\w\\s-_]+?)>");

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

    /*  *//**
     * 判断内容中是否含有指定的标签
     *
     * @return true：是，false：不是
     *//*
    public static String parseParamByTag(String content, String tagName) {
        // 构造正则
        String regStr = String.format("<%s>(.+)</%s>", tagName, tagName);
        // 只匹配第一个，因为我们只认第一个
        return ReUtils.get(regStr, content, 1);
    }

    public static String parseCellTagContent(Cell cell, String tagName) {
        if (cell == null) {
            return null;
        }

        // 获取单元格中的所有内容
        String cellValue = ExcelUtil.getCellValue(cell);
        // 获取指定标签的内容
        String parsedCellValue = pers.lzy.template.excel.common.TagParser.parseParamByTag(cellValue, tagName);
        // 如果是空，说明不是 simple-eval标签，则不处理
        if (StringUtils.isEmpty(parsedCellValue)) {
            return null;
        }
        return parsedCellValue;
    }

    public static List<String> parseTag(String content) {
        if (content == null) {
            return null;
        }
        return ReUtils.findAll(TAG_PATTERN, content, 2);
    }


    *//**
     * 获取内容中的第一个标签；
     *
     * @param content 被解析的内容
     * @return 解析出的第一个标签
     *//*
    public static String findFirstTag(String content) {
        if (content == null) {
            return null;
        }
        List<String> tagList = ReUtils.findAll(TAG_PATTERN, content, 2);
        if (tagList.size() == 0) {
            return null;
        }
        return tagList.get(0);
    }

    *//**
     * 获取内容中的第一个标签；
     *
     * @param cell 被解析的单元格
     * @return 解析出的第一个标签
     *//*
    public static String findFirstTag(Cell cell) {
        if (cell == null) {
            return null;
        }

        // 获取单元格中的所有内容
        String cellValue = ExcelUtil.getCellValue(cell);

        if (cellValue == null) {
            return null;
        }

        List<String> tagList = ReUtils.findAll(TAG_PATTERN, cellValue, 2);
        if (tagList.size() == 0) {
            return null;
        }
        return tagList.get(0);
    }
*/
}
