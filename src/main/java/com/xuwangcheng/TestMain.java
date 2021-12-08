package com.xuwangcheng;

import cn.hutool.core.io.FileUtil;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.xuwangcheng.html2word.HtmlToWordUtil;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 10:07
 */
public class TestMain {
    public static void main(String[] args) throws IOException {
        // 带表格
        //String html = FileUtil.readString("table_demo.html", "utf-8");
        // 包含图片
        String simpleText = FileUtil.readString("simple_text.html", "utf-8");
        String simpleTable01 = FileUtil.readString("simple_table_01.html", "utf-8");
        String simpleTable02 = FileUtil.readString("simple_table_02.html", "utf-8");
        String mixTable = FileUtil.readString("mix_table.html", "utf-8");
        String simpleImg = FileUtil.readString("simple_img.html", "utf-8");

        //配置
        Configure config = Configure.newBuilder().build();
        config.customPolicy("simpleText", HtmlToWordUtil.createHtmlRenderPolicy(null));
        config.customPolicy("simpleTable01", HtmlToWordUtil.createHtmlRenderPolicy(null));
        config.customPolicy("simpleTable02", HtmlToWordUtil.createHtmlRenderPolicy(null));
        config.customPolicy("mixTable", HtmlToWordUtil.createHtmlRenderPolicy(null));
        config.customPolicy("simpleImg", HtmlToWordUtil.createHtmlRenderPolicy(null));

        //创建word模板对象
        Map<String, Object> map = new HashMap<String, Object>();

        map.put("simpleText", simpleText);
        map.put("simpleTable01", simpleTable01);
        map.put("simpleTable02", simpleTable02);
        map.put("mixTable", mixTable);
        map.put("simpleImg", simpleImg);

        // 通用模板
        XWPFTemplate template = XWPFTemplate.compile(HtmlToWordUtil.getResourceInputStream("/out_template.docx"), config).render(map);
        template.writeToFile("I:\\out.docx");

        // 单元格内写入
        template = XWPFTemplate.compile(HtmlToWordUtil.getResourceInputStream("/table_cell_out_template.docx"), config).render(map);
        template.writeToFile("I:\\table_cell_out.docx");
        template.close();
    }
}
