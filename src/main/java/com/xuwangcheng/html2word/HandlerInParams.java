package com.xuwangcheng.html2word;

import com.deepoove.poi.NiceXWPFDocument;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jsoup.nodes.Element;

/**
 * 处理参数
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 10:15
 */
@Data
@Accessors(chain = true)
public class HandlerInParams {
    /**
     * 对应的html元素
     */
    private Element ele;
    /**
     * word段落
     */
    private XWPFParagraph xwpfParagraph;

    /**
     * 追加的起始块
     */
    private XWPFRun run;

    /**
     * 文档对象
     */
    private NiceXWPFDocument doc;

    /**
     * 其他可能需要的业务参数
     */
    private Object extendParams;
}
