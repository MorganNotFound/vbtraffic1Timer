# 产品信息 #
*产品名称：*`数码交通灯v2.0`   
*开发公司：*`港田路360号西马神奇动物园`   
*程序员：*`MorganFish`   
*邮箱：*
   `1037502595@qq.com`
   `MorganFish0508@163.com`
## CSDN:[http://blog.csdn.net/MorganFish](http://blog.csdn.net/MorganFish "MorganFish")
## GitHub:[https://www.github.com/MorganNotFound](https://www.github.com/MorganNotFound "MorganNotFound")
***以上信息与正文内容无关***   
# 代码 #
老规矩，感谢您的支持，话不多说，直接上代码：
```
Dim g As Integer
Dim y As Integer
Dim r As Integer
Private Sub cmdstart_Click()
cmdstart.Visible = False
cmdstop.Visible = True
Timer1.Enabled = True
Label1.Visible = True
Shape1.Visible = True
End Sub
Private Sub cmdstop_Click()
End
End Sub
Private Sub Form_Load()
Form1.BackColor = vbBlack
Label1.BackColor = vbBlack
Shape1.FillColor = vbGreen
Label1.ForeColor = vbGreen
Label1.Caption = g
Shape1.Visible = False
Label1.Visible = False
cmdstop.Visible = False
g = 10
y = 0
r = 0
Timer1.Enabled = False
End Sub
Private Sub Timer1_Timer()
Static n As Integer
   n = n + 1
   Shape1.BackStyle = 1
   Select Case n
      Case 1 To 10
         Shape1.BackColor = vbGreen
         Label1.Caption = g
         Label1.ForeColor = vbGreen
         g = g - 1
         y = 3
      Case 11 To 16
         If n Mod 2 = 1 Then
         Timer1.Interval = 500
            Shape1.BackColor = vbYellow
            Label1.Caption = y
            Label1.ForeColor = vbYellow
            y = y - 1
            r = 10
         Else
            Shape1.BackStyle = 0
         End If
      Case 17 To 26
      Timer1.Interval = 1000
         Shape1.BackColor = vbRed
         Label1.Caption = r
         Label1.ForeColor = vbRed
         r = r - 1
         g = 10
        If n = 26 Then n = 0
   End Select
End Sub
```
## 手把手教你制作： ##
（一）新建一个VisualBasic6.0工程，选择`标准exe`  
![](http://m.qpic.cn/psc?/cfc1fd56-f474-498f-adcb-b6fd8951402d/45NBuzDIW489QBoVep5mcSLU5m6SW6WoPMetn5hLcrL*TigNvaBdW5KxX*28TkQ.bUt4szPYtSLDnGLw1H.FRaK.yR0xQDe1yNc83SO6sW8!/b&bo=zQGaAQAAAAADF2U!&rf=viewer_4)    
（二）插入一个`shape`，一个`label`和一个`timer`，拖到合适位置并拖成合适大小    
（三）插入两个`CommandButton`控件，并将两个的`(名称)`属性分别改为`cmdstart`和`cmdstop`，注意，是名称，`Caption`自己随便取，并把两个拖到一起   
![](http://m.qpic.cn/psc?/cfc1fd56-f474-498f-adcb-b6fd8951402d/45NBuzDIW489QBoVep5mcbNK9Yx0dD*IZIl6fiA01TUrJAbepsm4ahzm*HqdU0M3KszEOj7.ssANJLX6eJfd*jLbTGSGy0mIE4Fpgypttxg!/b&bo=VgXPAgAAAAADF6w!&rf=viewer_4)   
（四）更改`Shape1.BackStyle`属性为`1 - Opaque`，并把`Shape1.Shape`属性改为`3 - Circle`   
（五）把`Label1.Font`自己改一下（建议初号）   
（六）将`Timer1.Interval`设定为1000   
（七）插入代码   
![](http://m.qpic.cn/psc?/cfc1fd56-f474-498f-adcb-b6fd8951402d/45NBuzDIW489QBoVep5mcfreLghjdDUXsVG4Ie9ME.zGIGmIh2N6rSSmRMBUHRRiUWnw.I03anUK*VZQ7z*Z1ICGFBhCvQ1Rbe5YHbixXMA!/b&bo=agOtAQAAAAADF*c!&rf=viewer_4)   
（八）运行   
![](http://m.qpic.cn/psc?/cfc1fd56-f474-498f-adcb-b6fd8951402d/45NBuzDIW489QBoVep5mcfreLghjdDUXsVG4Ie9ME.yHMFvWMD4cFezx0Wt1QOahhlrWSVDppyiGQpuB*6JJSd7kw.LncRQvorwxF2Lml1o!/b&bo=twHqAAAAAAADF24!&rf=viewer_4)   
![](http://m.qpic.cn/psc?/cfc1fd56-f474-498f-adcb-b6fd8951402d/45NBuzDIW489QBoVep5mcVuV0Vf2U9auphccIQtoAMqkzINIeq5RkC.4XqR0uu9u8aIGAYlYym.SpA*zSVRB7hI**EDdaXlkFcndDMAtZ4k!/b&bo=twHqAAAAAAADF24!&rf=viewer_4)   
![](http://m.qpic.cn/psc?/cfc1fd56-f474-498f-adcb-b6fd8951402d/45NBuzDIW489QBoVep5mcVyVZk7AkxzsXbcB.L8sPk1C0YY7SJfVIYUigtOyeibQzYNp*TfCashtzHpXheDJfx5RSSwqFut6C6L0RdioOw4!/b&bo=twHqAAAAAAADF24!&rf=viewer_4)   
![](http://m.qpic.cn/psc?/cfc1fd56-f474-498f-adcb-b6fd8951402d/45NBuzDIW489QBoVep5mcZwGoNE97zJz91WsRcU2YyZ*iRdha11s74tjdmnqKTzWznGOUiT8GSX8dVkE*yNUvT5lyFRyfbx7mXtw5ebfDEs!/b&bo=twHqAAAAAAADF24!&rf=viewer_4)   
# 工作原理 #   
1. 初始化程序   
```
Dim g As Integer
Dim y As Integer
Dim r As Integer
Private Sub form_load()
```   
2. 当单击`cmdstart`按钮，激活`Timer1`，然后隐藏`cmdstart`并显示`cmdstop`控件   
```
Private Sub cmdstart_Click()
cmdstart.Visible = False
cmdstop.Visible = True
Timer1.Enabled = True
Label1.Visible = True
Shape1.Visible = True
End Sub
```
3. 当`Timer1`被激活的时候，开始运行红绿灯；   
这一段中，我用了一个稍微超纲的一个指令`Select`…`Case`…`End Select`，并定义了一个`Integer` n，使用`n = n + 1`作为判断灯颜色的标识，`if`判断，进入循环。  
（中途为了使黄灯闪烁我调整了`Timer1.Interval = 500`然后又恢复为`1000`，影响不大）   
```
Private Sub Timer1_Timer()
Static n As Integer
   n = n + 1
   Shape1.BackStyle = 1
   Select Case n
      Case 1 To 10
         Shape1.BackColor = vbGreen
         Label1.Caption = g
         Label1.ForeColor = vbGreen
         g = g - 1
         y = 3
      Case 11 To 16
         If n Mod 2 = 1 Then
         Timer1.Interval = 500
            Shape1.BackColor = vbYellow
            Label1.Caption = y
            Label1.ForeColor = vbYellow
            y = y - 1
            r = 10
         Else
            Shape1.BackStyle = 0
         End If
      Case 17 To 26
      Timer1.Interval = 1000
         Shape1.BackColor = vbRed
         Label1.Caption = r
         Label1.ForeColor = vbRed
         r = r - 1
         g = 10
        If n = 26 Then n = 0
   End Select
End Sub
```
4. 最简单的一步，也很有用，方便调试：
```
Private Sub cmdstop_Click()
   End
End Sub
```
# 进阶 #
经高人指教，有关Label1的数码显示也有别的做法，可以简化程序：   
删除g,y,r三个整参数，然后将`Label.Caption = g/y/r`替换为`Abs(n - 11)`/`Abs(n - 17)`/`Abs(n - 27)`[也可以写成相反数，不用绝对值，`Interval`属性也可以不用更改，因本人未曾尝试，您自己可以自己调试，估计可以方便许多~]
# 附言 #
本人已将做完的成品放在文件夹内，可以直接调用，自己也可以修改并生成exe，而且有小惊喜哟~   
如果您喜欢我的作品，记得星标，关注并推荐给小伙伴哦   
---
***本文选自*** [https://github.com/MorganNotFound/vbtraffic1Timer](https://github.com/MorganNotFound/vbtraffic1Timer "MorganNotFound")   
***仅为个人使用，转载文章请注明出处，商业转载请联系作者***