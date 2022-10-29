package pers.lzy.template.word.test.function;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Before;
import org.junit.Test;
import pers.lzy.template.word.core.TemplateWordFiller;
import pers.lzy.template.word.core.TemplateWordFillerFactory;
import pers.lzy.template.word.test.pojo.ClassScore;
import pers.lzy.template.word.test.pojo.Thesis;
import pers.lzy.template.word.test.pojo.UserInfo;
import pers.lzy.template.word.utils.WordUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/7/12  9:19
 */
public class FunctionTest {


    private TemplateWordFiller templateWordFiller;

    @Before
    public void before() {

        // 构建导出者
        templateWordFiller = TemplateWordFillerFactory
                .defaultTemplateWordFillerBuilder()
                .expressionCacheSize(500)
                .build();

    }

    @Test
    public void test() throws IOException, InvalidFormatException {

        Map<String, Object> param = initParam();

        //获取文件的URL
        URL url = this.getClass()
                .getClassLoader()
                .getResource("score.docx");
        assert url != null;

        // 准备workbook
        Workbook workbook = new XSSFWorkbook();
        FileInputStream fileInputStream = new FileInputStream(new File(url.getPath()));
        XWPFDocument doc = new XWPFDocument(fileInputStream);


        // 调用我们的框架进行数据的填充
        templateWordFiller.fillData(() -> doc, () -> param);

        // 将Excel 写出去，查看结果
        WordUtil.writeFile(doc, new File("C:\\Users\\immort\\Desktop\\fff\\dddsscore.docx"));
    }

    /**
     * 初始化数据，可以是pojo类，或者是map等。
     *
     * @return 准备要填充的数据
     */
    private Map<String, Object> initParam() {


        Map<String, Object> res = new HashMap<>();
        ClassScore classScore = new ClassScore();
        res.put("classScore", classScore);

        classScore.setName("清华xx附属小学");
        classScore.setLevel("五年级一班");
        classScore.setPhone("15032000000");
        classScore.setTeacherName("zyliu-immort");


        List<ClassScore.Score> scoreList = new ArrayList<>();
        scoreList.add(new ClassScore.Score("张三", 80D, 98D, 30D));
        scoreList.add(new ClassScore.Score("李四", 70D, 88D, 88D));
        scoreList.add(new ClassScore.Score("王五", 90D, 61D, 90D));
        scoreList.add(new ClassScore.Score("赵六", 86D, 78D, 78D));

        classScore.setScore(scoreList);


        UserInfo userInfo = new UserInfo();
        userInfo.setAge(18);
        userInfo.setNickname("liuzy");
        userInfo.setAddress("中国北京市海淀区");
        userInfo.setTelPhone("173300300030");
        userInfo.setRemark("有车有房有钱，人生赢家");
        userInfo.setSex(1);


        res.put("userInfo", userInfo);

        Thesis thesis = buildThesis();

        res.put("thesis", thesis);
        return res;


    }

    private Thesis buildThesis() {
        Thesis thesis = new Thesis();
        thesis.setSubjectArgument("课题的研究界定重要包括三个方面：\n对国内外研究现状的分析材料一定要做好检索，查全资料、还要运用提纲挈领的语言来高度简明扼要的概括，要善于总结归纳。具体的说大体分步：一是要简要叙述收集到得资料的观点，二是要评论这些观点，有欠缺不深入的地方，三是针对这些缺欠的问题我要研究的观点是什么，也就是研究内容的深入部分。总体上可以说是“求全责备，一网打尽”。");
        thesis.setProjectDesign("方案设计是设计中的重要阶段，它是一个极富有创造性的设计阶段，同时也是一个十分复杂的问题，它涉及到设计者的知识水平、经验、灵感和想象力等。\n" +
                "SpringCloud\n" +
                "K8s\n" +
                "ReactJs");
        thesis.setProgressPlan("1、1- 3周：毕业实习、收集资料、撰写开题报告（3周）； 2、4周：工艺论证（1周）；\n" +
                "3、5-8周：工艺计算及设备设计、选型计算（4周）； 4、9-12周：绘图（4周）；\n" +
                "5、13周：技术经济分析及非工艺部分（1周）； 6、14：编制设计说明书（1周）；\n" +
                "7、15周：上交毕业设计说明书，答辩评审（1周）； 8、16周：毕业设计答辩（1周）。");
        thesis.setOpinionsInstructors("同意开题");
        thesis.setCommentsPanel("该生xxx牛牛，同意开题");
        return thesis;
    }
}
