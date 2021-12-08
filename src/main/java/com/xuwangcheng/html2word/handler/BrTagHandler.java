package com.xuwangcheng.html2word.handler;

import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * 换行符
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/12/8 16:25
 */
public class BrTagHandler extends BaseHtmlTagHandler {
    @Override
    public String getMatchTagName() {
        return "br";
    }

    @Override
    public HandlerOutParams handleHtmlElement(HandlerInParams params) {
        XWPFRun run = params.getXwpfParagraph().createRun();
        run.addBreak();

        return new HandlerOutParams();
    }
}
