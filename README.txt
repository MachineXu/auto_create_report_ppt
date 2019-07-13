使用：双击或直接执行dat2ppt.exe，执行过程无窗口

输入文件：
Input.dat              #程序输入文件，须与dat2ppt.exe同一目录
_temp_.pptx            #方案报告模板， 安装目录
_temp_comparison.pptx  #对比报告模板， 安装目录
image.png              #方案1报告图， 工作目录
image2.png             #方案2报告图， 工作目录
image3.png             #对比报告图1， 工作目录
image4.png             #对比报告图2， 工作目录
test1.txt              #方案1结果文件， 工作目录
test2.txt              #方案2结果文件， 工作目录

输出文件：
Output.dat             #程序执行结果或错误信息， 与dat2ppt.exe同一目录
report1.pptx           #方案1报告， 工作目录
report2.pptx           #方案1报告， 工作目录
report3.pptx           #对比报告， 工作目录


输入文件：Input.dat
格式1：(同时导出方案1报告、方案2报告、对比报告）
InstallPath = C:\Users\Qiang_Administrator\Desktop\pythonppt
WorkPath = C:\Users\Qiang_Administrator\Desktop\pythonppt
ImageName = image.png, image2.png, image3.png, image4.png
ReportName = report1.pptx, report2.pptx, report3.pptx
TxtName = test1.txt, test2.txt

格式2：(同时导出方案1报告、方案2报告）
InstallPath = C:\Users\Qiang_Administrator\Desktop\pythonppt
WorkPath = C:\Users\Qiang_Administrator\Desktop\pythonppt
ImageName = image.png, image2.png
ReportName = report1.pptx, report2.pptx
TxtName = test1.txt, test2.txt

格式3：(仅导出方案1报告）
InstallPath = C:\Users\Qiang_Administrator\Desktop\pythonppt
WorkPath = C:\Users\Qiang_Administrator\Desktop\pythonppt
ImageName = image.png
ReportName = report1.pptx
TxtName = test1.txt

注：
所有文件名必须有扩展名
多个文件名用逗号隔开
如果文件已存在则导出时覆盖同名文件

Output.dat内容及含义：
Succeed  #成功执行
NotFoundInputFile  #在本程序同目录下未发现Input.dat文件
[path/report.txt]正在使用中，请关闭后重试！  #当前ppt文件正在使用，无法导出，需要解除ppt占用后重新导出
