package com.xuwangcheng.html2word.handler;

import cn.hutool.core.util.ReUtil;
import cn.hutool.core.util.StrUtil;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * 标题
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/12/8 14:51
 */
public class TitleTagHandler extends BaseHtmlTagHandler {

    private static final int MAX_TITLE_FONT_SIZE = 28;

    @Override
    public String getMatchTagName() {
        return "h[12345]";
    }

    @Override
    public HandlerOutParams handleHtmlElement(HandlerInParams params) {
        // 取标题标签中所有的标题
        String text = params.getEle().wholeText();

        // 默认换行
        params.getXwpfParagraph().createRun().addBreak();

        if (StrUtil.isNotBlank(text)) {
            XWPFRun runText = params.getXwpfParagraph().createRun();
            // 设置标题样式
            Style headStyle = new Style();
            headStyle.setBold(true);
            headStyle.setFontSize(MAX_TITLE_FONT_SIZE - (ReUtil.getFirstNumber(params.getEle().tagName()) - 1) * 2);

            TextRenderPolicy.Helper.renderTextRun(runText, new TextRenderData(text, headStyle));
        }

        return new HandlerOutParams();
    }
}
