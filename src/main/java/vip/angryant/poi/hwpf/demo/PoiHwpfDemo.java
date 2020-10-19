package vip.angryant.poi.hwpf.demo;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterProperties;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class PoiHwpfDemo {
    public static void main(String[] args) throws Exception {
        Map<String, POIText> replaces = new HashMap<String, POIText>();
        replaces.put("${标题}", POIText.str("请假条"));
        replaces.put("${名称}", POIText.str("领导1"));
        replaces.put("${请假人}", POIText.str("爱因斯坦"));
        replaces.put("${请假时间}", POIText.str("2012-12-12"));

        // 读入模板文件并替换参数
        HWPFDocument hwpf = poiWordTableReplace("/Users/huteng/Desktop/template/templateTest.doc", replaces);

        // 另存模板文件为exportTestg.doc
        FileOutputStream out = new FileOutputStream("/Users/huteng/Desktop/template/exportTestg.doc");
        hwpf.write(out);
        out.flush();
        out.close();
    }

    public static HWPFDocument poiWordTableReplace(String sourceFile,
                                                   Map<String, POIText> replaces) throws Exception {
        FileInputStream in = new FileInputStream(sourceFile);
        HWPFDocument hwpf = new HWPFDocument(in);
        Range r = hwpf.getRange();
        CharacterProperties props = new CharacterProperties();
        props.setFontSize(10);

        for (int i = 0; i < r.numParagraphs(); i++) {
            Paragraph p = r.getParagraph(i);
            int numStyles = hwpf.getStyleSheet().numStyles();
            int styleIndex = p.getStyleIndex();
            if (numStyles > styleIndex) {
                String s = p.text();
                final String old = s;
                for (String key : replaces.keySet()) {
                    if (s.contains(key)) {
                        s = s.replace(key, replaces.get(key).getText());
                    }
                }
                if (!old.equals(s)) {// 有变化
                    p.replaceText(old, s);
                    s = p.text();
                    System.out.println("old:" + old + "->" + "s:" + s);
                }
            }
        }
        return hwpf;
    }
}
