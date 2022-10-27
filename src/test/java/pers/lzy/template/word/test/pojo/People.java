package pers.lzy.template.word.test.pojo;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/25  13:25
 */
public class People {

    private String name;
    private Integer age;
    private String remark;

    public People() {
    }

    public People(String name, Integer age, String remark) {
        this.name = name;
        this.age = age;
        this.remark = remark;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getRemark() {
        return remark;
    }

    public void setRemark(String remark) {
        this.remark = remark;
    }
}
