package com.xuwangcheng.html2word.handler;

import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 15:21
 */
public class SupTagHandler extends BaseHtmlTagHandler {
    @Override
    public String getMatchTagName() {
        return "sup";
    }

    @Override
    public HandlerOutParams handleHtmlElement(HandlerInParams params) {
        XWPFRun run = params.getXwpfParagraph().createRun();
        run.setText(params.getEle().text());
        // 设置字体加粗;
        run.setBold(true);
        // 设置字体大小;
        run.setFontSize(12);
        run.setFontFamily("Times New Roman", XWPFRun.FontCharRange.ascii);
        run.setFontFamily("宋体", XWPFRun.FontCharRange.eastAsia);
        run.setSubscript(VerticalAlign.SUPERSCRIPT);

        TextRenderPolicy.Helper.renderTextRun(run, new TextRenderData(params.getEle().text()));
        return new HandlerOutParams().setContinueItr(false);
    }
}
