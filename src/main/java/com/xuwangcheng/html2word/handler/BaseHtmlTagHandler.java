package com.xuwangcheng.html2word.handler;

import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;

/**
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 10:23
 */
public abstract class BaseHtmlTagHandler {

    /**
     *  匹配的html标签名称，如果匹配多个，请使用正则表达式，例如span|p
     * @author xuwangcheng
     * @date 2021/7/12 10:24
     * @param
     * @return {@link String}
     */
    public abstract String getMatchTagName();
    /**
     *  处理方式
     * @author xuwangcheng
     * @date 2021/7/12 10:28
     * @param params params
     * @return {@link HandlerOutParams}
     */
    public abstract HandlerOutParams handleHtmlElement(HandlerInParams params);
}
