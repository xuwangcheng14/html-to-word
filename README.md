## 更加完善的解决方案
> 大家可以参考这位朋友写的插件，基本上实现了大部分的html标签到word的转换：https://github.com/draco1023/poi-tl-ext

# Html To Word
### 富文本Html转Word

2019年的时候业务需要把前台的富文本html数据转换为WORD文档，百度了下，网上多数方案都是通过直接将整个html转换为web视图版的word,但是自己这边的需求是需要将很多很多个不同的富文本html嵌套组合到一起，并且还需要动态的增加标题，图片等内容，故WEB版的word方案舍弃。
后来百度发现了poi-tl（http://deepoove.com/poi-tl/）， 发现居然可以直接模版化word进行生成，反复研究了下个人的需求，最终采用的方案如下：
- 使用poi-tl模版化word,通poi-tl的嵌套模板功能动态生成文档标题及目录;
- 使用jfreeChart动态生成统计图片，通过poi-tl模板参数传入;
- 使用Jsoup解析富文本，通过poi-tl自定义策略将html各种标签转换为poi的word对象。


整理了下代码，希望能够给也有我这样需求的朋友提供下思路。


### 关于扩展标签解析处理
在com.xuwangcheng.html2word.handler包下新建对应的处理类，继承BaseHtmlTagHandler，实现getMatchTagName和handleHtmlElement方法即可。
关于具体的代码实现，需要你先要了解poi-tl的一些用法，参考http://deepoove.com/poi-tl/ 。

> 目前已经实现了，table,img,span,sup等标签，其他标签转换看个人需求自行实现，如果你有实现好的代码，欢迎PR，帮助更多的朋友。

### 有问题欢迎通过QQ，微信联系，联系方式在码云个人主页。