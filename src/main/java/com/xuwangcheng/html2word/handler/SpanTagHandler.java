package com.xuwangcheng.html2word.handler;

import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * 适用一些纯文本的标签
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 15:02
 */
public class SpanTagHandler extends BaseHtmlTagHandler {

    @Override
    public String getMatchTagName() {
        return "span";
    }

    @Override
    public HandlerOutParams handleHtmlElement(HandlerInParams params) {
        String text = params.getEle().wholeText();
        if (StringUtils.isNotBlank(text)) {
            XWPFRun run = params.getXwpfParagraph().createRun();
            TextRenderPolicy.Helper.renderTextRun(run, new TextRenderData(text));
        }

        return new HandlerOutParams().setContinueItr(false);
    }
}
