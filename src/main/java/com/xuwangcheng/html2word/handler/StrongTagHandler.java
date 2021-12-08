package com.xuwangcheng.html2word.handler;

import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * <b>加粗</b>
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/12/8 16:06
 */
public class StrongTagHandler extends BaseHtmlTagHandler {
    @Override
    public String getMatchTagName() {
        return "b|strong";
    }

    @Override
    public HandlerOutParams handleHtmlElement(HandlerInParams params) {
        String text = params.getEle().wholeText();
        if (StringUtils.isNotBlank(text)) {
            XWPFRun run = params.getXwpfParagraph().createRun();
            Style style = new Style();
            style.setBold(true);
            TextRenderPolicy.Helper.renderTextRun(run, new TextRenderData(text, style));
        }

        return new HandlerOutParams().setContinueItr(false);
    }
}
