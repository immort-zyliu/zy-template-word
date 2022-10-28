package pers.lzy.template.word.test.pojo;

import java.util.List;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/3/2  18:16
 * <p>
 * 班级成绩
 */
public class ClassScore {

    private String name;

    private String level;

    private String phone;

    private String teacherName;

    private List<Score> score;

    public String getTeacherName() {
        return teacherName;
    }

    public void setTeacherName(String teacherName) {
        this.teacherName = teacherName;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getLevel() {
        return level;
    }

    public void setLevel(String level) {
        this.level = level;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public List<Score> getScore() {
        return score;
    }

    public void setScore(List<Score> score) {
        this.score = score;
    }

    public static class Score {

        private String name;

        private Double chinese;

        private Double math;

        private Double english;

        public Score() {
        }

        public Score(String name, Double chinese, Double math, Double english) {
            this.name = name;
            this.chinese = chinese;
            this.math = math;
            this.english = english;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Double getChinese() {
            return chinese;
        }

        public void setChinese(Double chinese) {
            this.chinese = chinese;
        }

        public Double getMath() {
            return math;
        }

        public void setMath(Double math) {
            this.math = math;
        }

        public Double getEnglish() {
            return english;
        }

        public void setEnglish(Double english) {
            this.english = english;
        }
    }
}
