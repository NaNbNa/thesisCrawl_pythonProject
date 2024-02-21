## 基于python的可执行的GUI文献爬虫(知网,谷歌学术)
----------------------------
##### 文件和代码同统计
![统计](https://github.com/NaNbNa/thesisCrawl_pythonProject/assets/144761706/72b94425-d881-40bd-964d-f63985462e40)
1. 有学习用的代码,注释详细,已经从python打包成exe
2. 从两篇博客中受启发二写出的.给出他们的网址
3. https://blog.csdn.net/weixin_68789096/article/details/130900608
4. https://blog.csdn.net/bookssea/article/details/107309591
### 图形界面
![2](https://github.com/NaNbNa/thesisCrawl_pythonProject/assets/144761706/db3d1867-7b04-4e7e-911f-1b41829e2978)

### 功能介绍
1. 爬取知网镜像,谷歌学术镜像的文献(需要联网)
2. 选择目录/文件,可以将信息写入表格文件.
3. 目录是覆盖模式,文件是追加模式.覆盖--删除原有内容,写入新内容.追加--不删除原有内容,在末尾追加新内容..如果选择目录,文献信息会写入一个文件,其名为:搜索输入框的内容,后缀只支持.xlsx
4. 文献列表按照年份降序排序
5. 直接点击文献列表的一行,如果有链接,则会跳转链接
6. 每次爬取都是从网页的第一页开始
7. 可以中止爬取
8. 第一个文本框是爬取日志,记录运行情况
9. 第二个文本框是文献列表,可以快捷查看你想要的文献,并支持点击跳转
10. 两个文本框均可随时清空
### 效果展示
![image](https://github.com/NaNbNa/thesisCrawl_pythonProject/assets/144761706/f8553f32-fe71-40b9-a63e-068933d602ae)

