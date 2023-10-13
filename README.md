# SmHTTP.dll

神梦HTTP请求插件

<!-- PROJECT SHIELDS -->
[![VB6][VB6-shield]][VB6-url]
[![WinHTTP][WinHTTP-shield]][WinHTTP-url]
[![JSON][JSON-shield]][JSON-url]
[![httpbin][httpbin-shield]][httpbin-url]
[![Anjian][anjian-shield]][anjian-url]

<br />

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]

<!-- PROJECT LOGO -->
<br />

<p align="center">
  <a href="https://github.com/SMWHff/SmHTTP.dll/">
    <img src="logo.png" alt="Logo">
  </a>

  <h3 align="center">神梦HTTP请求插件</h3>
  <p align="center">
  支持GET、POST、HEAD等HTTP协议请求；支持构造请求协议头、请求Cookies；支持构造各种类型请求体（url、form、json）；支持解析JSON响应。希望这个插件能适用于开发者们的技术产品或服务与众不同的企业或个人。
    <br />
    <a href="https://github.com/SMWHff/SmHTTP.dll/Documents/SmHTTP.html"><strong>探索本项目的文档 »</strong></a>
    <br />
    <br />
    <a href="https://github.com/SMWHff/SmHTTP.dll/Examples">查看Demo</a>
    ·
    <a href="https://github.com/SMWHff/SmHTTP.dll/issues">报告Bug</a>
    ·
    <a href="https://github.com/SMWHff/SmHTTP.dll/issues">提出新特性</a>
  </p>

</p>


 本项目面向开发者
 
## 目录

- [上手指南](#上手指南)
  - [开发前的配置要求](#开发前的配置要求)
  - [安装步骤](#安装步骤)
- [文件目录说明](#文件目录说明)
- [开发的架构](#开发的架构)
- [部署](#部署)
- [使用到的框架](#使用到的框架)
- [贡献者](#贡献者)
  - [如何参与开源项目](#如何参与开源项目)
- [版本控制](#版本控制)
- [作者](#作者)
- [鸣谢](#鸣谢)

### 上手指南
1. 复制 `SmHTTP.dll`、`SmHTTP.html` 文件
2. 粘贴 到按键精灵根目录下的 `plugin` 文件夹里


###### 开发前的配置要求
1. 使用 Windows 操作系统
2. 安装 [Visual Basic 6.0][VB6-url]
3. 安装 [按键精灵2014][anjian-url]


###### **安装步骤**
1. Get a free API Key at [https://example.com](https://example.com)
2. Clone the repo
```sh
git clone https://github.com/SMWHff/SmHTTP.dll.git
```


### 文件目录说明

```
文件目录
├─bas/
├─cls/
├─Documents/
│  └─SmHTTP_chm/
│      ├─bin/
│      ├─css/
│      ├─html/
│      │  ├─其他命令/
│      │  ├─同步请求/
│      │  ├─基本命令/
│      │  ├─常见问题/
│      │  ├─异步请求/
│      │  └─构造命令/
│      └─js/
├─Examples/
│  ├─TC简单开发/
│  ├─VBScript/
│  └─按键精灵/
├─frm/
├─Releases/
├─res/
├─LICENSE
└─README.md
```


### 开发的架构 
![ARCHITECTURE](ARCHITECTURE.png)


### 部署
暂无


### 使用到的框架
- [WinHTTP](https://learn.microsoft.com/zh-cn/windows/win32/winhttp/using-winhttp)
- [JSON](https://www.json.org/json-zh.html)


### 贡献者
请阅读**CONTRIBUTING.md** 查阅为该项目做出贡献的开发者。


#### 如何参与开源项目
贡献使开源社区成为一个学习、激励和创造的绝佳场所。你所作的任何贡献都是**非常感谢**的。
1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request



### 版本控制
该项目使用Git进行版本管理。您可以在repository参看当前可用版本。


### 作者
1042207232@qq.com

昵称:神梦无痕  &ensp; qq:1042207232    

 *您也可以在贡献者名单中参看所有参与该项目的开发者。*



### 版权说明
该项目签署了 BSD 授权许可，详情请参阅 [LICENSE](https://github.com/SMWHff/SmHTTP.dll/blob/master/LICENSE)



### 鸣谢
- [GitHub Pages](https://pages.github.com)
- [GitHub Best_README_template](https://github.com/shaojintian/Best_README_template)
- [Choose an Open Source License](https://choosealicense.com)
- [Img Shields](https://shields.io)
- [JSON.org](https://www.json.org/)
- [按键精灵](https://www.anjian.com/)



<!-- links -->
[your-project-path]:SMWHff/SmHTTP.dll
[contributors-shield]: https://img.shields.io/github/contributors/SMWHff/SmHTTP.dll.svg?style=flat-square
[contributors-url]: https://github.com/SMWHff/SmHTTP.dll/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/SMWHff/SmHTTP.dll.svg?style=flat-square
[forks-url]: https://github.com/SMWHff/SmHTTP.dll/network/members
[stars-shield]: https://img.shields.io/github/stars/SMWHff/SmHTTP.dll.svg?style=flat-square
[stars-url]: https://github.com/SMWHff/SmHTTP.dll/stargazers
[issues-shield]: https://img.shields.io/github/issues/SMWHff/SmHTTP.dll.svg?style=flat-square
[issues-url]: https://img.shields.io/github/issues/SMWHff/SmHTTP.dll.svg
[license-shield]: https://img.shields.io/github/license/SMWHff/SmHTTP.dll.svg?style=flat-square
[license-url]: https://github.com/SMWHff/SmHTTP.dll/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=flat-square&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/SMWHff
[VB6-shield]: https://img.shields.io/badge/Visual%20Basic-6.0-blue?labelColor=512BD4
[VB6-url]: https://pc.qq.com/detail/19/detail_91139.html
[anjian-shield]: https://img.shields.io/badge/%E6%8C%89%E9%94%AE%E7%B2%BE%E7%81%B5-2014-white?logoColor=24ab5e&labelColor=24ab5e
[anjian-url]: http://www.anjian.com/
[WinHTTP-shield]: https://img.shields.io/badge/WinHTTP-5.1-blue
[WinHTTP-url]: https://learn.microsoft.com/zh-cn/windows/win32/winhttp/using-winhttp
[JSON-shield]: https://img.shields.io/badge/JSON-2009.4-blue
[JSON-url]: http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html
[httpbin-shield]: https://img.shields.io/badge/httpbin.org-0.9.2-blue
[httpbin-url]: https://httpbin.org/



