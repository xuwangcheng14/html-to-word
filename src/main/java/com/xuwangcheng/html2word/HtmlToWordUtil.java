package com.xuwangcheng.html2word;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.ClassUtil;
import cn.hutool.core.util.ReflectUtil;
import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.template.run.RunTemplate;
import com.xuwangcheng.html2word.handler.BaseHtmlTagHandler;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.InputStream;
import java.util.*;

/**
 * 工具类
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 10:09
 */
public class HtmlToWordUtil {


    private static final Map<String, BaseHtmlTagHandler> handlerMap = new HashMap<>();

    static {
        Set<Class<?>> classes = ClassUtil.scanPackageBySuper(ClassUtil.getPackage(BaseHtmlTagHandler.class), BaseHtmlTagHandler.class);
        if (CollUtil.isNotEmpty(classes)) {
            for (Class clazz:classes) {
                BaseHtmlTagHandler handler = (BaseHtmlTagHandler) ReflectUtil.newInstance(clazz);
                handlerMap.put(handler.getMatchTagName(), handler);
            }
        }
    }


    /**
     *  创建自定义处理逻辑
     * @author xuwangcheng
     * @date 2021/7/12 14:40
     * @param extendParams extendParams
     * @return {@link AbstractRenderPolicy}
     */
    public static AbstractRenderPolicy createHtmlRenderPolicy(Object extendParams) {
        return new AbstractRenderPolicy() {
            @Override
            protected void afterRender(RenderContext context) {
                // 清空模板标签所在段落
                clearPlaceholder(context, true);
            }

            @Override
            public void doRender(RunTemplate runTemplate, Object data, XWPFTemplate template) throws Exception {
                if (data == null || StringUtils.isBlank(data.toString())) {
                    return;
                }

                //获得Apache POI增强类NiceXWPFDocument
                NiceXWPFDocument doc = template.getXWPFDocument();


                String html = data.toString();
                // 处理html实体符号
                html = html.replaceAll("&gt;", ">")
                        .replaceAll("&lt;", "<")
                        .replaceAll("&nbsp;", " ")
                        .replaceAll("\\n", "")
                        .replaceAll("&crarr;", "")
                        .replaceAll("&quot;", "\"")
                        .replaceAll("&apos;", "'")
                        .replaceAll("&cent;", "￠")
                        .replaceAll("&pound;", "£")
                        .replaceAll("&yen;", "¥")
                        .replaceAll("&euro;", "€")
                        .replaceAll("&sect;", "§")
                        .replaceAll("&copy;", "©")
                        .replaceAll("&reg;", "®")
                        .replaceAll("&trade;", "™")
                        .replaceAll("&times;", "×")
                        .replaceAll("&divide;", "÷")
                        .replaceAll("&amp;", "&");
                org.jsoup.nodes.Document htmlDoc = Jsoup.parse(html);
                Elements nodes = htmlDoc.body().children();

                HandlerInParams inParams = new HandlerInParams();
                inParams.setEle(null);
                inParams.setExtendParams(extendParams);
                inParams.setXwpfParagraph(insertNewParagraph(((XWPFParagraph)runTemplate.getRun().getParent()).getBody(), doc));
                inParams.setRun(inParams.getXwpfParagraph().createRun());
                inParams.setDoc(doc);

                ListIterator<Element> itr = nodes.listIterator();
                while (itr.hasNext()) {
                    inParams.setEle(itr.next());
                    parseHtmlToWord(inParams);
                }
            }
        };
    }

    public static XWPFParagraph insertNewParagraph(IBody iBody, NiceXWPFDocument document){
        if (iBody instanceof XWPFTableCell) {
            XWPFTableCell tableCell = (XWPFTableCell) iBody;
            return tableCell.addParagraph();
        } else {
            return document.createParagraph();
        }
    }

    /**
     * 转换整个html内容为word内容
     * @author xuwangcheng
     * @date 2019/7/29 18:46
     * @param inParams inParams
     * @return {@link XWPFParagraph}
     */
    public static void parseHtmlToWord( HandlerInParams inParams) {
        Element ele = inParams.getEle();
        BaseHtmlTagHandler handler = getHandler(ele.tagName());
        HandlerOutParams outParams = handler.handleHtmlElement(inParams);
        if (outParams.getXwpfParagraph() != null) {
            inParams.setXwpfParagraph(outParams.getXwpfParagraph());
        }

        if (outParams.getContinueItr() && ele.children().size() > 0) {
            ListIterator<Element> itr = ele.children().listIterator();
            while (itr.hasNext()) {
                inParams.setEle(itr.next());
                parseHtmlToWord(inParams);
            }
        }

    }

    /**
     *  获取指定的标签处理器
     * @author xuwangcheng
     * @date 2021/7/12 14:40
     * @param tagName tagName
     * @return {@link BaseHtmlTagHandler}
     */
    private static BaseHtmlTagHandler getHandler(String tagName){
        if (tagName == null) {
            return null;
        }
        for (String key:handlerMap.keySet()) {
            if (key.equalsIgnoreCase(tagName) || tagName.toUpperCase().matches(key.toUpperCase())) {
                return handlerMap.get(key);
            }
        }

        return handlerMap.get("");
    }


    /**
     *  迭代获取指定段落
     * @author xuwangcheng
     * @date 2021/7/12 14:39
     * @param doc doc
     * @param xwpfParagraph xwpfParagraph
     * @return {@link XWPFParagraph}
     */
    public static XWPFParagraph getPrevXWPFParagraph (NiceXWPFDocument doc, XWPFParagraph xwpfParagraph) {
        List<XWPFParagraph> xwpfParagraphs = doc.getXWPFDocument().getParagraphs();
        for (int i = 0;i < xwpfParagraphs.size();i++) {
            if (xwpfParagraphs.get(i).equals(xwpfParagraph)) {
                return xwpfParagraphs.get(i + 1);
            }
        }

        return xwpfParagraph;
    }


    /**
     * 获取资源文件的文件流
     *
     * @return
     */
    public static InputStream getResourceInputStream(String filePath) {
        InputStream in = FileUtil.class.getResourceAsStream(filePath);
        if (in != null) {
            return in;
        }
        return null;
    }

}
