# coding=utf-8
# 用途：根据代码2生成的层次聚类结果调整订单批次，执行代码前需将"层次聚类结果.xlsx"文件中orderID一栏的格式手动从“2”“3”“4”修改为“ID000001”"ID000002""ID000003"等
# 输入："0305原始订单(去套单后-调顺序).xlsx" 、"层次聚类结果.xlsx"
# 输出："最终分批结果.xlsx"
# 最后一次更新时间：2022-10-06
# user：zhang heng
import pandas as pd
read_road01 = "D:/123/0305原始订单(去套单后-调顺序-含时间窗).xlsx"                       # 读取路径1
read_road02= "D:/123/层次聚类结果.xlsx"                                       # 读取路径2
df = pd.read_excel(read_road01)
batch = pd.read_excel(read_road02)
df.columns = ['BarCode','ProductID','Count','ShipCode','ShipTime']          # 导入文件,并更改列名
df['OderListID'] = None                                                     # 添加第六列，用于写入OderListID，每128个BarCode为一批
df = df[['OderListID','BarCode','ProductID','Count','ShipCode','ShipTime']] # 调整列的顺序，OderListID放在最前面

max_row = df.shape[0]                                   # 获取最大行数
for row in range(0,max_row):                            # 逐行判断订单号所属批次
    order = df.iloc[row, 1]                             # 获取该行订单号
    for i in range(0, 131):                             # 逐个与131个批次做对比
        batchID = batch.iloc[i, 0]                      # 获取第i行的批次号
        order_group = batch.iloc[i, 1]                  # 获取第i行批次中所含的订单号
        # 对字符串进行处理，去除'()空格4种无效字符
        order_group = order_group.replace("'", "")
        order_group = order_group.replace("(", "")
        order_group = order_group.replace(")", "")
        order_group = order_group.replace(" ", "")
        # 处理字符串END
        order_list = order_group.split(",")             # 将字符串按照逗号进行分割，保存为列表形式
        if order in order_list:                         # 若该行订单号位于第i行批次中
            df.iloc[row, 0] = batchID                   # 则将第i行的批次号写入该行的批次号
            break                                       # 已找到该订单的批次号，不必继续查找，跳出131个批次的查找循环，开始下一行订单的查找
write_road= "D:/123/最终分批结果.xlsx"                    # 写入路径
df.to_excel(write_road,index=False)                     # 写入文件，不加行号
print(df)