package pers.lzy.template.word.test.pojo;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/2/28  13:32
 */
public class PeopleProvince {

    private String provinceName;

    private String cityName;

    private String villageName;

    private String peopleName;


    public PeopleProvince() {
    }

    public PeopleProvince(String provinceName, String cityName, String villageName, String peopleName) {
        this.provinceName = provinceName;
        this.cityName = cityName;
        this.villageName = villageName;
        this.peopleName = peopleName;
    }

    public String getProvinceName() {
        return provinceName;
    }

    public void setProvinceName(String provinceName) {
        this.provinceName = provinceName;
    }

    public String getCityName() {
        return cityName;
    }

    public void setCityName(String cityName) {
        this.cityName = cityName;
    }

    public String getVillageName() {
        return villageName;
    }

    public void setVillageName(String villageName) {
        this.villageName = villageName;
    }

    public String getPeopleName() {
        return peopleName;
    }

    public void setPeopleName(String peopleName) {
        this.peopleName = peopleName;
    }
}
