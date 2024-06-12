# ChinaTobacco（合同组代码）

#### 介绍
烟草合同组代码仓库

#### 软件架构
软件架构说明
factors = {}                              //所有的要素
factors_ok = []                           //审查出正确的要素
factors_error = {}                         //审查出错误的要素
factors_to_inform = {}                     //提示信息
addRemarkInDoc(word, document, "甲方", "要素填写错误：甲方未填写完整")  //打批注信息

####文件说明
ContractType.py    判断文件类型
DocReader.py      读文件
Fuzzy_interface.py   
fuzzymatching.py
getFactorsFromContract.py 要素审查代码（其中包含了买卖合同和物业管理合同审查代码，为避免代码过多，房屋租赁审查代码写在了rent_contract.py中，采购合同代码写在了purchase_and_warehousing_contract.py中，在getFactorsFromContract.py中写了接口）

helpful.py 和 utils.py 里面写了要素审查用到的各种函数

main.py 
None_standard_contract.py
over_all_description.py

UiLayout.py         界面代码
Userdict.txt          jieba分词库

sbert.pkl 这是利用sbert进行模糊匹配的模型直接复制过去（模型过大，已经放到网盘中，下载链接：https://pan.baidu.com/s/1cAI7u6nTgohDPnjc3fwm4g?pwd=1111 
提取码：1111）


#### 安装教程

1.  运行main.py
2.  选择待审查文件
3.  选择输出路径
4.  开始审查

#### 安装包注意事项

1.  需要有sbert.skl(网盘链接：https://pan.baidu.com/s/1cAI7u6nTgohDPnjc3fwm4g?pwd=1111 
提取码：1111)
2.  安装所有提示的库
3.  text2vec包只能用0.2.2版本的，其他版本会报错

# 安装包及其命令 add by qy Date:2022.5.16
1. pip install PyQt5
2. pip install python-docx
3. pip install pywin32
4. pip install numpy
5. pip install fuzzywuzzy
6. pip3 install torch torchvision torchaudio
7. pip install sentence_transformers
8. pip install text2vec==0.2.2
9. pip install six
10. pip install xlwt

# Sbert 模型添加方案  add by qy Date:2022.5.18
修改自动下载模型，不使用编译保存的模型即可：  
print("加载Sbert………………")   
device = torch.device("cuda" if torch.cuda.is_available() else "cpu")   
print("运行设备：", device)   
model = SBert('paraphrase-multilingual-MiniLM-L12-v2').to(device)  
print("加载Sbert完成…………"   

使用缓存模型方法：
sentence-transformers_paraphrase-multilingual-MiniLM-L12-v2.zip  （链接：提取码：）  
解压添加至：       
C:\Users\用户名\.cache\torch\sentence_transformers\sentence-transformers_paraphrase-multilingual-MiniLM-L12-v2

# 封装API add by qy Date:2022.5.18
文件： Hetong_api.py  
命令： python Hetong_api.py file1 file2 file3 export_path 
