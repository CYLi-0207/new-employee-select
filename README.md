# 本月有哪些新员工入职？

## 分析过程

生成一段代码，确保其按照如下步骤处理输入的数据：
1.检查字段“入职日期”范围是否在2025-03-01到2025-03-31之间，若是则保留，若否则排除
2.检查字段“员工二级类别”字段是否是正式员工，若是则保留，若否则排除
3.检查字段“四级组织”下的值是否“证照支持部”，若是则排除，若否则保留
4.特殊情况：杨晶文（“员工系统号”的值为“31049588”）、李云飞（“员工系统号”的值为“31268163”）属于特殊情况，如果命中则排除
5.确保遍历所有的数据，分别标记保留和排除
6.对于符合条件的记录，提取如下字段的值：“三级组织”、“姓名”、“花名”、“入职日期”、“员工二级类别”。
7.输出结果按照三级组织的值由大到小排序，每个三级组织下按照入职日期的值由小到大排序
8.实现上述规则后，运行代码并返回保留人员明细

生成第二段代码，分析上一段代码输出的保留人员明细
1.统计并输出返回了多少条数据
2.将上述的返回结果进行加工，生成两列输出，第一列是返回结果中“三级组织”的唯一值，第二列是返回结果中“三级组织”与唯一值匹配的所有人员的信息拼接，拼接格式为“姓名（花名）”，如有多个符合条件的值，用“、”分隔
3.提醒用户检查：特殊原因外包人员、活水人员

实现上述规则后，运行代码并返回完整的人员信息拼接结果

## 前端实现

仔细阅读以上的代码，仔细理解并一步一步列出代码表示的输入、每一步计算过程和输出（共2个文件）

根据上述代码中的输入、输出和计算过程，形成一个完整的可在streamlit部署的代码文档，实现如下的功能
1.在开始计算前，页面上显示如下的内容：（1）页面上方的说明：“本网页根据2025.4.4版本的花名册数据生成，如果输入数据有变更，产出可能出错，需要与管理员联系”；（2）页面下方标记两个输出文件的下载框，但在分析完成之前不能下载，分析完成后才可下载
2.导入数据：界面功能需要支持导入花名册数据，每次只能导入一份  
3.页面需要支持选择年份和月份，根据选择的年份和月份确定输出的人员信息中，入职日期的范围应当在哪年哪月范围内
4.点击“开始分析”后，执行上述代码的每一步分析过程，分步骤产出两个输出文档，每个输出文档产出后，对应的文件可以在前端下载
5.错误反馈：如果导入的数据和系统要求的字段不符，返回对应的错误说明
6.计算完成后显示的提醒中需包括：因特殊规则剔除的人员；需要额外考虑的人员（活水入职）
