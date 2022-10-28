package pers.lzy.template.word.test.function.func;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/7/12  9:20
 */
public class ScoreUtil {

    /**
     * 分数格式化方法
     *
     * @param score 分数
     * @return 格式化后的结果
     */
    public static String formatScore(double score) {

        String evaluation;

        if (score >= 90) {
            evaluation = "优秀";
        } else if (score >= 80) {
            evaluation = "良好";
        } else if (score >= 60) {
            evaluation = "尚可";
        } else {
            evaluation = "不及格";
        }
        return evaluation;
    }
}
