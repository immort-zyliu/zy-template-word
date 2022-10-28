package pers.lzy.template.word.test.simple;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pers.lzy.template.word.core.TemplateWordFillerFactory;
import pers.lzy.template.word.filler.DefaultTemplateWordFiller;
import pers.lzy.template.word.test.pojo.PeopleProvince;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static pers.lzy.template.word.utils.WordUtil.writeFile;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/27  17:55
 */
public class SimpleTest {

    public static void main(String[] args) throws Exception {
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\immort\\Desktop\\fff\\11.docx");
        XWPFDocument doc = new XWPFDocument(fileInputStream);


        DefaultTemplateWordFiller.Builder builder = TemplateWordFillerFactory.defaultTemplateWordFillerBuilder();

        DefaultTemplateWordFiller defaultTemplateWordFiller = builder.build();

        doFill(defaultTemplateWordFiller, doc);


        writeFile(doc, new File("C:\\Users\\immort\\Desktop\\fff\\1111111-test.docx"));
    }

    private static void doFill(DefaultTemplateWordFiller defaultTemplateWordFiller, XWPFDocument doc) {

        defaultTemplateWordFiller.fillData(doc, initParam());
    }


    private static Map<String, Object> initParam() {
        Map<String, Object> res = new HashMap<>();

        // 联系人信息
        Map<String, Object> contractInfo = new HashMap<>();
        contractInfo.put("contactName", "liuzy");
        contractInfo.put("contactTel", "15020000000");
        res.put("contractInfo", contractInfo);


        // 人员信息（注意顺序的问题，如果顺序不对，是不能进行合并的。。。）
        List<PeopleProvince> peopleProvinceList = new ArrayList<>();
        peopleProvinceList.add(new PeopleProvince("河北省", "石家庄市", "兼顾村", "张三"));
        peopleProvinceList.add(new PeopleProvince("河北省", "石家庄市", "兼顾村", "张三"));
        peopleProvinceList.add(new PeopleProvince("河北省", "保定市", "牛村", "李四2"));
        peopleProvinceList.add(new PeopleProvince("河北省", "保定市", "牛村", "李四3"));
        peopleProvinceList.add(new PeopleProvince("北京市", "海淀区", "厉害村", "风"));
        peopleProvinceList.add(new PeopleProvince("北京市", "朝阳区", "牛逼村", "扽给"));
        res.put("peopleInfo", peopleProvinceList);

        return res;
    }
}
