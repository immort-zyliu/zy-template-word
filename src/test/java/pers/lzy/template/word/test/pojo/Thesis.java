package pers.lzy.template.word.test.pojo;

import java.awt.*;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/29  16:09
 * <p>
 * 论文对象
 */
public class Thesis {

    /**
     * 系
     */
    private String department = "计算机";

    /**
     * 专业
     */
    private String major = "计算机科学与技术专业";

    /**
     * 年级（那一届的？）
     */
    private Integer grade = 2017;

    /**
     * 学生姓名
     */
    private String stuName = "immort-zyliu";


    /**
     * 论文题目
     */
    private String thesisTitle = "基于SpringCloud的多租户云xx系统的设计与实现";

    /**
     * 指导教师
     */
    private String tutorName = "郝教授";


    /**
     * 职称
     */
    private String technicalLevel = "教授";


    /**
     * 指导教师 所属系别
     */
    private String tutorDept = "计算机系";

    /**
     * 研究方向
     */
    private String researchDirection = "操作系统内核";


    /**
     * 课题论证
     */
    private String subjectArgument;

    /**
     * 方案设计
     */
    private String projectDesign;

    /**
     * 进度计划
     */
    private String progressPlan;

    /**
     * 指导教师意见
     */
    private String opinionsInstructors;

    /**
     * 指导小组意见
     */
    private String commentsPanel;

    /**
     * 指导教师签字
     */
    private String tutorSign = "郝教授";

    /**
     * 组长签字
     */
    private String groupLeaderSign = "高教授";

    private String year = "2021";

    private String month = "11";

    private String days = "12";

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }

    public String getMajor() {
        return major;
    }

    public void setMajor(String major) {
        this.major = major;
    }

    public Integer getGrade() {
        return grade;
    }

    public void setGrade(Integer grade) {
        this.grade = grade;
    }

    public String getStuName() {
        return stuName;
    }

    public void setStuName(String stuName) {
        this.stuName = stuName;
    }

    public String getThesisTitle() {
        return thesisTitle;
    }

    public void setThesisTitle(String thesisTitle) {
        this.thesisTitle = thesisTitle;
    }

    public String getTutorName() {
        return tutorName;
    }

    public void setTutorName(String tutorName) {
        this.tutorName = tutorName;
    }

    public String getTechnicalLevel() {
        return technicalLevel;
    }

    public void setTechnicalLevel(String technicalLevel) {
        this.technicalLevel = technicalLevel;
    }

    public String getTutorDept() {
        return tutorDept;
    }

    public void setTutorDept(String tutorDept) {
        this.tutorDept = tutorDept;
    }

    public String getResearchDirection() {
        return researchDirection;
    }

    public void setResearchDirection(String researchDirection) {
        this.researchDirection = researchDirection;
    }

    public String getSubjectArgument() {
        return subjectArgument;
    }

    public void setSubjectArgument(String subjectArgument) {
        this.subjectArgument = subjectArgument;
    }

    public String getProjectDesign() {
        return projectDesign;
    }

    public void setProjectDesign(String projectDesign) {
        this.projectDesign = projectDesign;
    }

    public String getProgressPlan() {
        return progressPlan;
    }

    public void setProgressPlan(String progressPlan) {
        this.progressPlan = progressPlan;
    }

    public String getOpinionsInstructors() {
        return opinionsInstructors;
    }

    public void setOpinionsInstructors(String opinionsInstructors) {
        this.opinionsInstructors = opinionsInstructors;
    }

    public String getCommentsPanel() {
        return commentsPanel;
    }

    public void setCommentsPanel(String commentsPanel) {
        this.commentsPanel = commentsPanel;
    }

    public String getTutorSign() {
        return tutorSign;
    }

    public void setTutorSign(String tutorSign) {
        this.tutorSign = tutorSign;
    }

    public String getGroupLeaderSign() {
        return groupLeaderSign;
    }

    public void setGroupLeaderSign(String groupLeaderSign) {
        this.groupLeaderSign = groupLeaderSign;
    }

    public String getYear() {
        return year;
    }

    public void setYear(String year) {
        this.year = year;
    }

    public String getMonth() {
        return month;
    }

    public void setMonth(String month) {
        this.month = month;
    }

    public String getDays() {
        return days;
    }

    public void setDays(String days) {
        this.days = days;
    }
}
