# 疫情期间excel分类统计自动化
为了帮助广大老师简化工作流程,特地学了几天的PYTHON,水平不高,多多包涵
范例数据全部虚拟化,保护隐私
## 开发环境安装
1. 安装python3.8 
    - [Download](https://github.com/3038922/new_century_robotics/releases/download/v1.0/python-3.8.1-amd64.exe)
2. 安装openpyxl 
    - 以管理员身份打开 `powershell`
    - 执行 `pip install openpyxl`

## 使用说明
1. 首先用钉钉收集班主任统计来的源数据 推荐番茄表单
2. 大致是这么制作.如果否就不提示后续,如果有再跳出要填的表格减轻班主任负担
![avatar](./pic/1.jpg)
3. 下载本项目,弄个文件夹存放
    - [选另存为](https://github.com/3038922/SARS-CoV-2_autoFill/blob/master/main.py)
4. 把班主任提交的源数据 放进同名文件夹.
5. 根据实际情况自己修改
![avatar](./pic/2.jpg)
6. 命令行里执行 `python3.8 ./main.py` 