package com.xuwangcheng.html2word;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

/**
 * 处理出参
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 10:16
 */
@Data
@Accessors(chain = true)
public class HandlerOutParams {
    /**
     * 是否不再迭代处理子节点数据
     */
    private Boolean continueItr = true;
    /**
     * 返回新的word段落，如果没有，则继续在老的段落后追加
     */
    private XWPFParagraph xwpfParagraph;
}
