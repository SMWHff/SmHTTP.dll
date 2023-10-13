//#================================================================
//#         神梦HTTP请求插件 SmHTTP.dll 演示 TC 例子
//#----------------------------------------------------------------
//#        【作者】：神梦无痕
//#        【ＱＱ】：1042207232
//#        【Ｑ群】：624655641
//#        【更新】：2020-10-17
//#----------------------------------------------------------------
//#  插件说明：用于HTTP协议的请求访问操作
//#----------------------------------------------------------------
//#  神梦工具：http://pan.baidu.com/s/1dESHf8X
//#================================================================


变量 线程ID = 0
变量 运行状态


function 神梦HTTP请求例子()
    
    如果(注册神梦HTTP插件())
        traceprint("**********************【神梦HTTP请求插件 SmHTTP.dll 演示 TC 例子】**********************")
        
        // 定义变量
        traceprint("定义变量")
        var user, pass, Data, Ret, Cookies, Headers
        
        // 配置账号
        user = 文件读配置("配置", "user", "D:\\Main.ini") // 你的按键精灵论坛账号
        pass = 文件读配置("配置", "pass", "D:\\Main.ini") // 你的按键精灵论坛密码
        Data = strformat("username=%s&password=%s", user, pass)
        
        // 开启自动识别参数模式
        SmHTTP.SetAutoParamArray(true)
        
        
        // 登录论坛账号
        Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/login.aspx?referer=forumindex.aspx", Data)
        // 判断是否登录成功
        if(InStr(Ret, user) == 0)  
            MsgBox("出错，登录失败！", "报错！")
            return
        end
        Cookies = SmHTTP.GetCookies()
        
        // 打卡签到
        Data = strformat("signmessage=%s", "签个到，每天心情都是美美哒~~按键精灵祝大家新年好运连连！！")
        Headers = strformat("Referer:%s", "http://bbs.anjian.com/")
        Ret = SmHTTP.HTTP_POST("http://bbs.anjian.com/addsignin.aspx?infloat=1&inajax=1", Data, Headers, Cookies)
        if(InStr(Ret, "恭喜您获取本日签到奖励") > 0 || InStr(Ret, "你今天已经签到过了") > 0) // 判断是否签到成功
            traceprint("恭喜，您已完成签到任务！")
        end
        
        
        
        
        // 构造JSON请求体
		//{
		//	"query": {
		//		"status": {
		//			"option": "online"
		//		},
		//		"type": "神圣石",
		//		"stats": [{
		//			"type": "and",
		//			"filters": []
		//		}],
		//		"filters": {
		//			"trade_filters": {
		//				"filters": {
		//					"price": {
		//						"option": "chaos"
		//					}
		//				}
		//			}
		//		}
		//	},
		//	"sort": {
		//		"price": "asc"
		//	}
		//}
        var jsonData = array(), query = array()
        query["status"] = array("option"="online")
        query["type"] = "神圣石"
        query["stats"] = array(array("type"="and", "filters"=array()))
        query["filters"] = array("trade_filters"=array("filters"=array("price"=array("option"="chaos"))))
        
        jsonData["query"] = query
        jsonData["sort"] = array("price"="asc")
        
        traceprint(ArrayToJSON(jsonData))
    结束
    MsgBox("脚本执行完毕!")
end



//注册神梦HTTP插件
function 注册神梦HTTP插件()
    如果(文件是否存在("rc:SmHTTP.dll") == 假)
        消息框("注册失败, 请先把 SmHTTP.dll 插件\n放到【神梦HTTP请求示例】目录下的【资源】文件夹里!")
        返回 假
    结束
    变量 Ret = 注册插件("rc:SmHTTP.dll", 真)
    如果(Ret)
        调试输出("注册成功")
    否则
        消息框("注册失败,请尝试其他方式注册")
        返回 假
    结束
    SmHTTP = 插件("SMWH.SmHTTP")
    如果(获取变量类型(SmHTTP) != "com")
        消息框("没有注册插件")
        返回 假
    结束
    调试输出("插件版本号: " & SmHTTP.Ver())
    返回 真
end

function 线程状态检测()
    循环(线程获取状态(线程ID))
        等待(100)
    结束
    蜂鸣(1200, 90)
    线程ID = 0
    调试输出("脚本已经停止执行")
    控件是否有效("启动按钮", 真)
end
//启动_热键操作
function 启动_onhotkey()
    蜂鸣(900, 200)
    如果(线程ID == 0)
        控件是否有效("启动按钮", 假)
        线程ID = 线程开启("神梦HTTP请求例子", "")
        线程开启("线程状态检测", "")
    否则
        消息框("脚本正在执行中,请先停止再启动!")
    结束
end
function 启动按钮_click()
    启动_onhotkey()
end


//终止热键操作
function 终止_onhotkey()
    蜂鸣(1100, 90)
    如果(线程ID != 0)
        线程关闭(线程ID)
    结束    
end
function 终止按钮_click()
    终止_onhotkey()
end


function 启动_killfocus()
    //这里添加你要执行的代码
    热键销毁("启动")
    热键注册("启动")
end


function 终止_killfocus()
    //这里添加你要执行的代码
    热键销毁("终止")
    热键注册("终止")
end


function 保存配置_click()
    //这里添加你要执行的代码
    变量 键值 = 0, 功能键 = 0
    热键获取键码("启动", 键值, 功能键)
    文件写配置("热键", "启动键值", 键值, "D:\\Main.ini")
    文件写配置("热键", "启动功能键", 功能键, "D:\\Main.ini")
    
    热键获取键码("终止", 键值, 功能键)
    文件写配置("热键", "终止键值", 键值, "D:\\Main.ini")
    文件写配置("热键", "终止功能键", 功能键, "D:\\Main.ini")
end


function 神梦工具_click()
    //下载神梦抓抓工具
    命令("http://pan.baidu.com/s/1dESHf8X", false)
end

function 神梦HTTP请求示例_init()
    //窗体初始化事件
    变量 键值 = 0, 功能键 = 0
    键值 = 文件读配置("热键", "启动键值", "D:\\Main.ini")
    功能键 = 文件读配置("热键", "启动功能键", "D:\\Main.ini")
    如果(键值 != "")
        热键设置键码("启动", 键值, 功能键)
        热键注册("启动")
    结束
    
    键值 = 文件读配置("热键", "终止键值", "D:\\Main.ini")
    功能键 = 文件读配置("热键", "终止功能键", "D:\\Main.ini")
    如果(键值 != "")
        热键设置键码("终止", 键值, 功能键)
        热键注册("终止")
    结束
    
    变量 提示内容 = "【神梦HTTP请求插件 SmHTTP.dll 演示TC例子】\n\n"
    提示内容 = 提示内容 & "支持多种请求类型\n\n"
    提示内容 = 提示内容 & "支持异步请求方式\n\n"
    标签设置文本("标签2", 提示内容)
    
    提示内容 = "作者：神梦无痕\n\n"
    提示内容 = 提示内容 & "ＱＱ：1042207232\n\n"
    提示内容 = 提示内容 & "Ｑ群：624655641\n\n"
    标签设置文本("标签3", 提示内容)
    
    
    //设置托盘
    变量 标题 = 窗口获取标题(窗口获取自我句柄()) 
    设置托盘(标题, 假)
end


