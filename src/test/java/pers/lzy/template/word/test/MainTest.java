package pers.lzy.template.word.test;

import com.alibaba.fastjson.JSONObject;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pers.lzy.template.word.core.TemplateWordFillerFactory;
import pers.lzy.template.word.filler.DefaultTemplateWordFiller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.HashMap;

import static pers.lzy.template.word.utils.WordUtil.writeFile;

/**
 * @author immort-liuzyj(zyliu)
 * @since 2022/10/27  17:32
 */
public class MainTest {

    public static void main(String[] args) throws Exception {

        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\immort\\Desktop\\fff\\test.docx");
        XWPFDocument doc = new XWPFDocument(fileInputStream);


        DefaultTemplateWordFiller.Builder builder = TemplateWordFillerFactory.defaultTemplateWordFillerBuilder();

        DefaultTemplateWordFiller defaultTemplateWordFiller = builder.build();

        defaultTemplateWordFiller.fillData(doc, new HashMap<>());

        writeFile(doc, new File("C:\\Users\\immort\\Desktop\\fff\\123132.docx"));
    }
}
