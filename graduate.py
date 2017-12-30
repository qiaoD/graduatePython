'''
Author：QiaoD
Time：2017/12/25

研究生管理系统简化：
1、课表查询
2、自动评教

UI : Tkinter
http : requests
Html : BeautifulSoup

'''


import requests
import tkinter as tk
from bs4 import *
from PIL import Image, ImageTk
import tkinter.messagebox as Box

class Graduate:

    # 域名
    domain = "http://202.114.200.86/"
    # 登录地址
    loginUrl = "UserLogin.aspx?exit=1"
    # 评教地址
    TeachEvaluateUrl = "gstudent/Message/TeachEvaluate.aspx?EID=9Wzxl!at!rH89FB!iBcZKvBD79WFmupS-QrWraU4VOlyww7zS!Y88g==&msg=1&UID="
    TeachEvalGradeUrl = "Gstudent/Message/TeachEvalGrade.aspx"
    # 成绩查询地址
    ScoreQueryUrl = "Gstudent/Course/StudentScoreQuery.aspx?EID=g99JT9GMEz32OaY0-RKDUnJuxn2BPYuPmMdNO483f!wb5tCKe8yCNg==&UID="

    name = ""                                   # 学号
    cookies = {"LoginType" : "LoginType=1",}    # 登录cookies
    loginState = 0                              # 登录状态


    # login
    def __init__(self):
        loginPage = self.domain + self.loginUrl
        re = requests.get(loginPage)
        text = re.text                  # 解析登录页
        #print(text)
        self.loginPost = {
            "__EVENTTARGET" : "btLogin",
            "__EVENTARGUMENT" : "",
            "__LASTFOCUS" : "",
            "__EVENTVALIDATION" : "",                       # 获取
            "__VIEWSTATE" : "",                             # 获取
            "ScriptManager1" : "UpdatePanel2|btLogin",
            "UserName" : "",
            "PassWord" : "",
            "ValidateCode" : "",
            "drpLoginType" : "1",
            "__ASYNCPOST" : "true",
        }
        soup = BeautifulSoup(text, "lxml")
        ValidateImg = ""
        for tag in soup.findAll(name = "input"):
            if "__EVENTVALIDATION" == tag.attrs["name"] :
                self.loginPost['__EVENTVALIDATION'] = tag.attrs['value']
            if "__VIEWSTATE" == tag.attrs["name"]:
                self.loginPost['__VIEWSTATE'] = tag.attrs['value']
            if "ValidateImage" == tag.attrs["name"]:
                ValidateImg = self.domain + tag.attrs['src']

        # 获取验证码，同时获取cookies
        imgtext = requests.get(ValidateImg)

        with open("ValidateCode.jpg","wb") as f:
            f.write(imgtext.content)
        f.close()
        self.cookies["ASP.NET_SessionId"] = imgtext.cookies["ASP.NET_SessionId"]

    def login(self,UserName="",PassWord="",ValidateCode=""):
        loginPage = self.domain + self.loginUrl
        self.loginPost["UserName"] = UserName
        self.loginPost["PassWord"] = PassWord
        self.loginPost["ValidateCode"] = ValidateCode

        self.name = self.loginPost["UserName"]
        # login
        headers = {
            "Host": "202.114.200.86",
            "Connection": "keep-alive",
            "Origin": "http://202.114.200.86",
            "X-Requested-With": "XMLHttpRequest",
            "Cache-Control": "no-cache",
            "X-MicrosoftAjax": "Delta=true",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Accept": "*/*",
            "Referer": "http://202.114.200.86/UserLogin.aspx?exit=1",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Cookie": "ASP.NET_SessionId="+self.cookies["ASP.NET_SessionId"]+"; LoginType=LoginType=1;",
        }


        login = requests.post(url = loginPage, data = self.loginPost, cookies = self.cookies, headers = headers)
        #print(login.text)
        if(-1 != login.text.find("pageRedirect")):
            self.loginState = 1


    def TeachEvaluate(self):

        TeachEvaluateUrl = self.domain + self.TeachEvaluateUrl + self.name
        res = requests.get(url = TeachEvaluateUrl,cookies = self.cookies)
        # print(res.text)
        soup = BeautifulSoup(res.text, "lxml")
        table = soup.select('table[class="Grid_Line"]')
        trs = table[0].findAll(name="tr")
        #print(trs)
        del(trs[0])         # 去掉表格头部
        evaluates = []
        for tr in trs:
            #print(tr)
            evaluate = {}
            tds = tr.findAll(name="td")
            #print(len(tds))
            #evaluate["dep"]         = tds[0].string
            #evaluate["number"]      = tds[-10].string
            # evaluate["lessonName"]  = str(tds[-9].string).strip()
            evaluate["teacher"]     = str(tds[-8].span.string).strip()
            #evaluate["beginTime"]   = tds[-7].span.string
            #evaluate["endTime"]     = tds[-6].span.string
            #evaluate["status"]      = tds[-5].span.string
            evaluate["score"]       = tds[-4].span.string
            evaluate["isEnd"]       = tds[-3].span.string
            #evaluate["date"]        = tds[-2].span.string
            evaluate["link"]        = self.domain + self.TeachEvalGradeUrl + tds[-1].a.attrs["onclick"][8:-12]
            #if(not evaluate["score"]):
            evaluates.append(evaluate)
        # 开始评教 表单
        sum = 0
        for ev in evaluates:
            #print("正在为" + ev["teacher"] + "的《" + ev["lessonName"] + "》课程进行评教-")
            url = ev['link']
            #print(url)
            #s = input("ss:")
            cookies = self.cookies
            score = 90
            data = {
                "__VIEWSTATEENCRYPTED" : "",
                "__EVENTVALIDATION" : "",
                "__EVENTTARGET" : "ctl00$contentParent$lnkSave",
                "__EVENTARGUMENT" : "",
                "__VIEWSTATE" : "",
                "ctl00$ScriptManager1" : "ctl00$ScriptManager1|ctl00$contentParent$lnkSave",
                "ctl00$contentParent$txtPjyj" : "",
                "ctl00$contentParent$dgData$ctl02$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl03$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl04$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl05$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl06$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl07$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl08$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl09$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl10$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl11$drpPfdj" : score,
                "ctl00$contentParent$dgData$ctl12$drpPfdj" : score,
                "__ASYNCPOST" : "true",
            }
            re = requests.get(url = url,cookies = cookies)
            sp = BeautifulSoup(re.text, "lxml")
            for tag in sp.findAll(name = "input"):
                if "__VIEWSTATEENCRYPTED" == tag.attrs["name"] :
                    data['__VIEWSTATEENCRYPTED'] = tag.attrs['value']
                if "__EVENTVALIDATION" == tag.attrs["name"]:
                    data['__EVENTVALIDATION'] = tag.attrs['value']
                if "__VIEWSTATE" == tag.attrs["name"]:
                    data['__VIEWSTATE'] = tag.attrs['value']

            headers = {
                "Host": "202.114.200.86",
                "Connection": "keep-alive",
                "Origin": "http://202.114.200.86",
                "X-Requested-With": "XMLHttpRequest",
                "Cache-Control": "no-cache",
                "X-MicrosoftAjax": "Delta=true",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "Accept": "*/*",
                "Referer": TeachEvaluateUrl,
                "Accept-Encoding": "gzip, deflate",
                "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
                "Cookie": "ASP.NET_SessionId="+self.cookies["ASP.NET_SessionId"]+"; LoginType=LoginType=1;",
            }

            rep = requests.post(url = url, data = data, cookies = cookies, headers = headers)
            #if 200 == rep.status_code:
                #print(ev["teacher"] + "的《" + ev["lessonName"] + "》课程评教成功！！！")
            sum += 1
            item = {}
            item["teacher"] = ev["teacher"]
            #item["lessonName"] = ev["lessonName"]
            yield item


    def StudentScoreQuery(self):
        ScoreQueryUrl = self.domain + self.ScoreQueryUrl + self.name
        res = requests.get(url = ScoreQueryUrl,cookies = self.cookies)
        soup = BeautifulSoup(res.text, "lxml")
        table = soup.select('table[class="Grid_Line"]')
        trs = table[0].findAll(name="tr")
        del(trs[0])         # 去掉表格头部
        item = {}
        for tr in trs:
            tds = tr.findAll(name="td")
            item["name"] = tds[0].string
            item["time"] = tds[1].string
            item["xuefen"] = tds[2].string
            item["term"] = tds[3].string
            item["kaoqin"] = tds[4].string
            item["work"] = tds[5].string
            item["exam"] = tds[6].string
            item["score"] = tds[7].string
            item["rank"] = tds[8].string
            item["type"] = tds[9].string
            item["shuxing"] = tds[10].string
            item["beizhu"] = tds[11].string
        yield item


'''
G = Graduate()
G.TeachEvaluate()
'''
class Application(tk.Frame):
    def __init__(self, master=None, G = None):
        super().__init__(master)
        self.G = G
        self.master.title("Hole -V1")
        #self.master.iconbitmap("hole.ico")
        self.master.resizable(False, False)
        self.master.geometry("500x350")
        #self.pack()
        self.login()

    def login(self):
        self.loginFrame = tk.Frame(self.master,width=500,height=350,bg = "white")
        self.loginFrame.place(x=0,y=0)

        # title
        self.TitleLabel = tk.Label(self.loginFrame, width=20, text="南望山职业技术学院",fg="#336699", bg="white" )
        self.TitleLabel.place(x=175,y=40)

        # UserName
        self.UserNameLabel = tk.Label(self.loginFrame, width=8, text="学号：", bg="white")
        self.UserNameLabel.place(x=100,y=100)
        self.UserNameEntry = tk.Entry(self.loginFrame, width=25, bd=1)
        self.UserNameEntry.place(x=200,y=100)

        # pwd
        self.PwdLabel = tk.Label(self.loginFrame, width=8,text="密码：", bg="white")
        self.PwdLabel.place(x=100,y=130)
        self.PwdEntry = tk.Entry(self.loginFrame, width=25,  bd=1, show = "*")
        self.PwdEntry.place(x=200,y=130)

        # ValidateImage
        self.ValidLabel = tk.Label(self.loginFrame, width=8,text="验证码：", bg="white")
        self.ValidLabel.place(x=100,y=160)
        self.ValidEntry = tk.Entry(self.loginFrame, width=15,  bd=1)
        self.ValidEntry.place(x=200,y=160)
        # ValidateCode.jpg

        #load the jpg
        load = Image.open('ValidateCode.jpg') #
        render= ImageTk.PhotoImage(load)
        self.ValidImgLabel = tk.Label(self.loginFrame, bg="white", image = render)
        self.ValidImgLabel.image = render
        self.ValidImgLabel.place(x=320,y=160)

        # loginButton
        self.LoginButton = tk.Button(self.loginFrame, text = "登录",width=8,command = self.loginPost)
        self.LoginButton.place(x=120,y=210)

        # @
        self.TitleLabel = tk.Label(self.loginFrame, width=20, text="@wherer.cn",fg="#336699", bg="white" )
        self.TitleLabel.place(x=175,y=300)


    def loginPost(self):
        UserName = self.UserNameEntry.get()
        Pwd      = self.PwdEntry.get()
        Valid    = self.ValidEntry.get()
        self.G.login(UserName=UserName, PassWord=Pwd, ValidateCode=Valid)
        if(1 == self.G.loginState):
            Box.showinfo("success","欢迎您！！！")
            self.loginFrame.destroy()
            self.dash()
        else :
            Box.showinfo("error","登陆失败！！！")



    def dash(self):
        self.dashFrame = tk.Frame(self.master,width=500,height=350,bg = "white")
        self.dashFrame.place(x=0,y=0)
        self.EvaluateButton = tk.Button(self.dashFrame, text = "一键评教",width=10,command = self.RunTeachEvaluate)
        self.EvaluateButton.place(x=120,y=50)
        self.QueryButton = tk.Button(self.dashFrame, text = "成绩查询",width=10,command = self.RunStudentScoreQuery)
        self.QueryButton.place(x=290,y=50)
        #显示区域
        self.var = tk.StringVar(self.dashFrame)
        self.showText = tk.Message(self.dashFrame, textvariable=self.var, width=300,bg="white")
        self.showText.place(x=50,y=100)

    def RunTeachEvaluate(self):
        self.var.set("")
        for item in self.G.TeachEvaluate():
            tmp = self.var.get()
            self.var.set(tmp + item["teacher"] + "评教成功！！！\n")

    def RunStudentScoreQuery(self):
        self.var.set("")
        for item in self.G.StudentScoreQuery():
            tmp = self.var.get()
            self.var.set(tmp + item["name"] + ":" + item["score"] + "\n")

if __name__ == '__main__'
G = Graduate()
root = tk.Tk()
app = Application(master=root,G=G)

app.mainloop()
