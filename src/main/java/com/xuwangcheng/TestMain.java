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
        String html = FileUtil.readString("img_demo.html", "utf-8");;
        //配置
        Configure config = Configure.newBuilder().build();
        config.customPolicy("resultHtml", HtmlToWordUtil.createHtmlRenderPolicy(null));


        //创建word模板对象
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("top", "TOPPPPP");
        map.put("resultHtml", html);
        map.put("buttom", "buttommmmmmmmmmmmmm");
        XWPFTemplate template = XWPFTemplate.compile(HtmlToWordUtil.getResourceInputStream("/out_template.docx"), config).render(map);
        template.writeToFile("I:\\demo.docx");
        template.close();
    }
}
