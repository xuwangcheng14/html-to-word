package com.xuwangcheng.html2word.handler;

import cn.hutool.core.io.FileUtil;
import cn.hutool.http.HttpUtil;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.policy.PictureRenderPolicy;
import com.xuwangcheng.html2word.HandlerInParams;
import com.xuwangcheng.html2word.HandlerOutParams;
import lombok.SneakyThrows;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.jsoup.nodes.Element;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.UUID;

/**
 * @author xuwangcheng
 * @version 1.0.0
 * @description
 * @date 2021/7/12 15:08
 */
public class ImgTagHandler extends BaseHtmlTagHandler {
    @Override
    public String getMatchTagName() {
        return "img";
    }

    @SneakyThrows
    @Override
    public HandlerOutParams handleHtmlElement(HandlerInParams params) {
        //获取图片本地路径
        String imgRealPath = getImgRealPath(params.getEle());

        if (StringUtils.isBlank(imgRealPath) || !FileUtil.exist(imgRealPath)) {
            return new HandlerOutParams().setContinueItr(false);
        }

        //插入图片
        //获取图片对象
        BufferedImage img = ImageIO.read(new File(imgRealPath));

        //获得图片的宽
        int width = img.getWidth();
        //获得图片的高
        int height = img.getHeight();
        if (width > 600) {
            //获取比例
            int rate = (width / 600 ) + 1;
            width = width / rate - 20;
            height = height / rate;
        }
        PictureRenderData pictureRenderData = new PictureRenderData(width, height, imgRealPath);
        XWPFRun run = params.getXwpfParagraph().createRun();
        PictureRenderPolicy.Helper.renderPicture(run, pictureRenderData);
        return new HandlerOutParams().setContinueItr(false);
    }


    /**
     * TODO 这个方法是通过img标签的属性去获取图片文件，需要根据业务来编写
     * @author xuwangcheng
     * @date 2019/11/21 9:47
     * @param ele ele
     * @return {@link String}
     */
    private String getImgRealPath (Element ele) {
        String imgUrl = ele.attr("src");
        if (imgUrl.startsWith("data:;base64,")) {
            // 如果是base64需要转换下
            imgUrl = imgUrl.replace("data:;base64,", "");
            // base64转本地图片
            byte[] bytes = Base64.getDecoder().decode(imgUrl);
            String path = FileUtil.getTmpDirPath() + File.separator + UUID.randomUUID().toString() + ".png";
            copyByte2File(bytes, new File(path));
            return path;
        } else if (imgUrl.matches("^(?i)(http|https)://.*$")) {
            // 通过网络下载图片
            String path = FileUtil.getTmpDirPath() + File.separator + UUID.randomUUID().toString() + ".png";
            HttpUtil.downloadFile(imgUrl, path);
            return path;
        } else {
            // TODO 通过其他方式，比如数据库查询

        }
        return null;
    }

    private boolean copyByte2File(byte [] bytes,File file){
        FileOutputStream out = null;
        try {
            //转化为输入流
            ByteArrayInputStream in = new ByteArrayInputStream(bytes);

            //写出文件
            byte[] buffer = new byte[1024];

            out = new FileOutputStream(file);

            //写文件
            int len = 0;
            while ((len = in.read(buffer)) != -1) {
                out.write(buffer, 0, len);
            }
            return true;
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
            try {
                if(out != null){
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return false;
    }
}
