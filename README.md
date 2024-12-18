## @超链接函数
```
=HYPERLINK("D:\file\"&A2, A2)
```
## 多列内容合并
```
=OFFSET($A$2,(ROW(A1)-1)/2,MOD(ROW(A1)-1,2))&"" 两列转1列工具，交插存放文本
```
## 一个工作薄多个sheet页合并-（适用于office自带excel，WPS不可用）
按Alt+F11两键，
调出Visual Basic 界面，
在左侧窗口中，右键选择“插入”—“模块”，将下列代码粘贴进去，点击运行按钮，完成数据表合并。
Sub 合并当前工作簿下的所有工作表()
```
On Error Resume Next
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set st = Worksheets.Add(before:=Sheets(1))
st.Name = "合并"
For Each shet In Sheets:
If shet.Name <> "合并" Then
i = st.Range("A" & Rows.Count).End(xlUp).Row + 1
shet.UsedRange.Copy
st.Cells(i, 1).PasteSpecial Paste:=xlPasteAll
End If
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "已完成"
End Sub
```

如果数据是在一列上，A1：A10的则可在A11输入：
对奇数行求和：
```
=SUMPRODUCT((MOD(ROW(A1:A10),2)=1)*A1:A10)
```
对偶数行求和：
```
=SUMPRODUCT((MOD(ROW(A1:A10),2)=0)*A1:A10)
```
如果数据是在一行上，如A1：J1，则可在K1输入：
对奇数列求和：
```
=SUMPRODUCT((MOD(COLUMN(A1:J1),2)=1)*A1:J1)
```
对偶数列求和：
```
=SUMPRODUCT((MOD(COLUMN(A1:J1),2)=0)*A1:J1)
```

## 把所有wav复制到输出路径：
```
for /r 输入绝对路径 %i in (*.wav) do copy %i 输出路径
```
## 把所有wav剪切到输出路径：
```
for /r 输入绝对路径 %i in (*.wav) do move %i 输出路径
```
## 把所有txt文本合并到输出路径：
```
for /r 输入绝对路径 %i in (*.txt) do (type "%i" >>输出绝对路径\result.txt)
```
```
eg:
for /r C:\Users\xiangli65\Desktop\HaitianCreole句式重复  %i in (*.txt) do (type "%i" >>C:\Users\xiangli65\Desktop\HaitianCreole句式重复\output\result.txt)
```
把所有txt文件复制到输出路径：
```
for /r 输入绝对路径 %i in (*.txt) do copy %i 指定目录
```

## 批量重命名：
第一步：
进入需要重命名的文件夹中，打开命令行【cmd】，输入命令
```
dir/b>rename.csv
```
第二步：

打开rename.csv文件。可以看见第一列就是我们当前文件夹下的所有文件名，然后在第二列输入每一个需要重命名后的名字

第三步：

在第三列输入公式
```【="ren "&""""&A1&""""&" "&B1】```，生成重命名字符串 
注意：在连接单元格时，如果该单元格有特殊符号时需要加上双引号（也就是需要输入四个双引号）没特殊符不需要加双引号

第四步：

复制第三列内容到当前文件夹下新建的txt文件中，重命名为bat文件

第五步：

双击运行bat文件，得到重命名文件
