package com.xuwangcheng.html2word.handler;

import cn.hutool.core.util.ReUtil;
import cn.hutool.core.util.StrUtil;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 16:26
 */
public class CommonTextTagHandler extends BaseHtmlTagHandler {
    @Override
    public String getMatchTagName() {
        return "";
    }

    @Override
    public HandlerOutParams handleHtmlElement(HandlerInParams params) {
        // 默认取自身标签中的文字，而不包括子节点
        String text = params.getEle().ownText();

        boolean enabledBreak = ReUtil.isMatch("(p|h[12345]|li|img)", params.getEle().tagName());
        if (enabledBreak) {
            XWPFRun run = params.getXwpfParagraph().createRun();
            run.addBreak();
        }

        if (StrUtil.isNotBlank(text)) {
            XWPFRun run = params.getXwpfParagraph().createRun();
            TextRenderPolicy.Helper.renderTextRun(run, new TextRenderData(text));
        }

        return new HandlerOutParams();
    }
}
