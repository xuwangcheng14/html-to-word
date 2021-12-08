package com.xuwangcheng.html2word.handler;

import cn.hutool.core.bean.BeanUtil;
import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.util.TableTools;
import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;
import com.xuwangcheng.html2word.HtmlToWordUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.select.Elements;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;

import java.util.*;

/**
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 14:41
 */
public class TableTagHandler extends BaseHtmlTagHandler {

    @Override
    public String getMatchTagName() {
        return "table";
    }

    @Override
    public HandlerOutParams handleHtmlElement(HandlerInParams params) {
        IBody iBody = params.getXwpfParagraph().getBody();
        XWPFParagraph nextParagraph = null;
        if (iBody instanceof XWPFTableCell) {
            XWPFTableCell tableCell = (XWPFTableCell) iBody;
            params.setXwpfParagraph(tableCell.addParagraph());
            params.setRun(params.getXwpfParagraph().createRun());
            parseTableToWord(params);
            nextParagraph = tableCell.addParagraph();
        } else {
            params.setXwpfParagraph(params.getDoc().createParagraph());
            params.setRun(params.getXwpfParagraph().createRun());
            parseTableToWord(params);
            nextParagraph = params.getDoc().createParagraph();
        }

        return new HandlerOutParams().setXwpfParagraph(nextParagraph).setContinueItr(false);
    }

    /**
     * 转换表格为word内容
     * @author xuwangcheng
     * @date 2019/7/29 18:45
     * @param inParams inParams
     */
    private void parseTableToWord(HandlerInParams inParams) {
        Element ele = inParams.getEle();
        NiceXWPFDocument doc = inParams.getDoc();
        XWPFRun run = inParams.getRun();

        //简化表格html
        org.jsoup.nodes.Document tableDoc = Jsoup.parse(simplifyTable(ele.outerHtml()));
        Elements trList = tableDoc.getElementsByTag("tr");
        Elements tdList = trList.get(0).getElementsByTag("td");

        //创建表格
        XWPFTable xwpfTable = doc.insertNewTable(run, trList.size(), tdList.size());
        CTTblLayoutType type = xwpfTable.getCTTbl().getTblPr().addNewTblLayout();
        type.setType(STTblLayoutType.AUTOFIT);


        //设置样式
        TableTools.widthTable(xwpfTable, MiniTableRenderData.WIDTH_A4_FULL, tdList.size());
        TableTools.borderTable(xwpfTable, 4);

        //写入表格行和列内容
        Map<String, Boolean>[][] array = new Map[trList.size()][tdList.size()];
        for (int row = 0; row < trList.size(); row++) {
            Element trElement = trList.get(row);
            Elements tds = trElement.getElementsByTag("td");
            for (int col = 0; col < tds.size(); col++) {
                Element colElement = tds.get(col);
                String colspan = colElement.attr("colspan");
                String rowspan = colElement.attr("rowspan");
                String style = colElement.attr("style");
                StringBuilder styleSB = new StringBuilder();
                if (!StringUtils.isEmpty(colspan)) {
                    int colCount = Integer.parseInt(colspan);
                    for (int i = 0; i < colCount - 1; i++) {
                        array[row][col + i + 1] = new HashMap<String, Boolean>();
                        array[row][col + i + 1].put("mergeCol", true);
                    }
                }
                if (!StringUtils.isEmpty(rowspan)) {
                    int rowCount = Integer.parseInt(rowspan);
                    for (int i = 0; i < rowCount - 1; i++) {
                        array[row + i + 1][col] = new HashMap<String, Boolean>();
                        array[row + i + 1][col].put("mergeRow", true);
                    }
                }
                XWPFTableCell tableCell = xwpfTable.getRow(row).getCell(col);
                if (StringUtils.isEmpty(colspan)) {
                    if (col == 0) {
                        if (tableCell.getCTTc().getTcPr() == null) {
                            tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                        } else {
                            if (tableCell.getCTTc().getTcPr().getHMerge() == null) {
                                tableCell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.RESTART);
                            } else {
                                tableCell.getCTTc().getTcPr().getHMerge().setVal(STMerge.RESTART);
                            }
                        }
                    } else {
                        if (array[row][col] != null && array[row][col].get("mergeCol") != null && array[row][col].get("mergeCol")) {
                            if (tableCell.getCTTc().getTcPr() == null) {
                                tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                            } else {
                                if (tableCell.getCTTc().getTcPr().getHMerge() == null) {
                                    tableCell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
                                } else {
                                    tableCell.getCTTc().getTcPr().getHMerge().setVal(STMerge.CONTINUE);
                                }
                            }
                            continue;
                        } else {
                            if (tableCell.getCTTc().getTcPr() == null) {
                                tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                            } else {
                                if (tableCell.getCTTc().getTcPr().getHMerge() == null) {
                                    tableCell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.RESTART);
                                } else {
                                    tableCell.getCTTc().getTcPr().getHMerge().setVal(STMerge.RESTART);
                                }
                            }
                        }
                    }
                } else {
                    if (tableCell.getCTTc().getTcPr() == null) {
                        tableCell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
                    } else {
                        if (tableCell.getCTTc().getTcPr().getHMerge() == null) {
                            tableCell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.RESTART);
                        } else {
                            tableCell.getCTTc().getTcPr().getHMerge().setVal(STMerge.RESTART);
                        }
                    }
                }
                if (StringUtils.isEmpty(rowspan)) {
                    if (array[row][col] != null && array[row][col].get("mergeRow") != null && array[row][col].get("mergeRow")) {
                        if (tableCell.getCTTc().getTcPr() == null) {
                            tableCell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
                        } else {
                            if (tableCell.getCTTc().getTcPr().getVMerge() == null) {
                                tableCell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
                            } else {
                                tableCell.getCTTc().getTcPr().getVMerge().setVal(STMerge.CONTINUE);
                            }
                        }
                        continue;
                    } else {
                        if (tableCell.getCTTc().getTcPr() == null) {
                            tableCell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
                        } else {
                            if (tableCell.getCTTc().getTcPr().getVMerge() == null) {
                                tableCell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.RESTART);
                            } else {
                                tableCell.getCTTc().getTcPr().getVMerge().setVal(STMerge.RESTART);
                            }
                        }
                    }
                } else {
                    if (tableCell.getCTTc().getTcPr() == null) {
                        tableCell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
                    } else {
                        if (tableCell.getCTTc().getTcPr().getVMerge() == null) {
                            tableCell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.RESTART);
                        } else {
                            tableCell.getCTTc().getTcPr().getVMerge().setVal(STMerge.RESTART);
                        }
                    }
                }
                tableCell.removeParagraph(0);
                XWPFParagraph paragraph = tableCell.addParagraph();
                paragraph.setStyle(styleSB.toString());
                if (!StringUtils.isEmpty(style) && style.contains("text-align:center")) {
                    paragraph.setAlignment(ParagraphAlignment.CENTER);
                }

                HandlerInParams handlerInParams = new HandlerInParams();
                BeanUtil.copyProperties(inParams, handlerInParams);

                handlerInParams.setEle(colElement);
                handlerInParams.setXwpfParagraph(paragraph);
                handlerInParams.setRun(paragraph.createRun());
                HtmlToWordUtil.parseHtmlToWord(handlerInParams);
            }
        }
    }

    /**
     * 简化html中的表格dom
     * @author xuwangcheng
     * @date 2019/7/29 18:39
     * @param tableContent tableContent
     * @return {@link String}
     */
    private String simplifyTable(String tableContent) {
        if (StringUtils.isEmpty(tableContent)) {
            return null;
        }
        org.jsoup.nodes.Document tableDoc = Jsoup.parse(tableContent);
        Elements trElements = tableDoc.getElementsByTag("tr");
        if (trElements != null) {
            Iterator<Element> eleIterator = trElements.iterator();
            Integer rowNum = 0;
            // 针对于colspan操作
            while (eleIterator.hasNext()) {
                rowNum++;
                Element trElement = eleIterator.next();
                //去除所有样式
                trElement.removeAttr("class");
                Elements tdElements = trElement.getElementsByTag("td");
                List<Element> tdEleList = covertElements2List(tdElements);
                for (int i = 0; i < tdEleList.size(); i++) {
                    Element curTdElement = tdEleList.get(i);
                    //去除所有样式
                    curTdElement.removeAttr("class");
                    Element ele = curTdElement.clone();
                    String colspanValStr = curTdElement.attr("colspan");
                    if (!StringUtils.isEmpty(colspanValStr)) {
                        ele.removeAttr("colspan");
                        Integer colspanVal = Integer.parseInt(colspanValStr);
                        for (int k = 0; k < colspanVal - 1; k++) {
                            curTdElement.after(ele.outerHtml());
                        }
                    }
                }
            }
            // 针对于rowspan操作
            List<Element> trEleList = covertElements2List(trElements);
            Element firstTrEle = trElements.first();
            Elements tdElements = firstTrEle.getElementsByTag("td");
            Integer tdCount = tdElements.size();
            //获取该列下所有单元格
            for (int i = 0; i < tdElements.size(); i++) {
                for (Element trElement : trEleList) {
                    List<Element> tdElementList = covertElements2List(trElement.getElementsByTag("td"));
                    try {
                        tdElementList.get(i);
                    } catch (Exception e) {
                        continue;
                    }
                    Node curTdNode = tdElementList.get(i);
                    Node cNode = curTdNode.clone();
                    String rowspanValStr = curTdNode.attr("rowspan");
                    if (!StringUtils.isEmpty(rowspanValStr)) {
                        cNode.removeAttr("rowspan");
                        Element nextTrElement = trElement.nextElementSibling();
                        Integer rowspanVal = Integer.parseInt(rowspanValStr);
                        for (int j = 0; j < rowspanVal - 1; j++) {
                            Node tempNode = cNode.clone();
                            List<Node> nodeList = new ArrayList<Node>();
                            nodeList.add(tempNode);
                            if (j > 0) {
                                nextTrElement = nextTrElement.nextElementSibling();
                            }
                            Integer indexNum = i + 1;
                            if (i == 0) {
                                indexNum = 0;
                            }
                            if (indexNum.equals(tdCount)) {
                                nextTrElement.appendChild(tempNode);
                            } else {
                                nextTrElement.insertChildren(indexNum, nodeList);
                            }
                        }
                    }
                }
            }
        }
        Element tableEle = tableDoc.getElementsByTag("table").first();
        String tableHtml = tableEle.outerHtml();

        return tableHtml;
    }

    /**
     * 转换Elements为list
     * @author xuwangcheng
     * @date 2019/7/29 18:40
     * @param curElements curElements
     * @return {@link List}
     */
    private List<Element> covertElements2List(Elements curElements){
        List<Element> elementList = new ArrayList<Element>();
        Iterator<Element> eleIterator = curElements.iterator();
        while(eleIterator.hasNext()){
            Element curlement = eleIterator.next();
            elementList.add(curlement);
        }
        return elementList;
    }
}
