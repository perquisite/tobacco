合同放进data文件夹
执行InspectionOfContract.py文件即可

--2021-08-07--
各代码解释如下：
        yancaoRegularDemo各代码功能为：
        Main.py：主程序，调动界面与后端
        ReadFile.py：读取所要处理的docx文件的信息
        remakefile.py：！！！此程序不参与发布版本！！！其作用为将【data】文件夹中的文件，转化为可供处理的docx文件，并存入【副本】文件夹
        tangyuhao.py:表1到表18的处理入口
        xiejunyu.py:表19到表36的处理入口
        UiLayout.py：window.ui使用QtDesigner转化而成的代码文件，供Main.py调用
        window.ui：界面文件，可使用QtDesigner打开与修改
---------------------------------------------------------------------------
此次版本更新后，需对以下部分进行修改：
        Main.py：54行代码d = Precessor2(file, self.export_dir)是用于处理表19到表36，如需处理表1到表18，需修改tangyuhao.py
                 后修改此处，并增加import

        tangyuhao.py: 请参考xiejunyu.py 进行修改，以适配界面

        table_1_18 与 table_19_36：文件夹中的所有表，请参考表19进行修改。修改点如下：
                                  1.增加了返回信息的info_list,将提示信息append到info_list中
                                  2.修改了_display函数
                                  3.修改check函数，返回info_list
                                  4.修改了init函数的传入参数，增加了self.source_prifix，用于在表中打开其它表时使用

--2021-08-07--
此次版本更新后，需对以下部分进行修改：
        table_1_18 与 table_19_36：文件夹中的所有表，请参考表19进行修改。修改点如下：
                                    1.增加了check函数中，对func的异常处理


--2021-10-08--
初步完成需求一打案件对比批注
每套文书中必须要有 案件处理审批表、行政处罚决定书、立案报告表、结案报告表
批注打在了 案件处理审批表 中