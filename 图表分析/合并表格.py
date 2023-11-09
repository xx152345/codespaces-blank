import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import time
from matplotlib.ticker import FuncFormatter, ScalarFormatter
import matplotlib as mpl

plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
mpl.rcParams['font.size'] = 14# 设置全局字体大小
plt.rcParams['figure.figsize']=(12.8, 7.2) # 全局设置输出图片大小 1280 x 720 像素
plt.ticklabel_format(style='plain',scilimits=(0,0),axis='both')#关闭科学计数法

def merge_excel_files(folder_path, output_file):
    global combined_data  # 声明combined_data为全局变量
    # 初始化一个空的DataFrame，用于存储合并后的数据
    combined_data = pd.DataFrame()

    # 遍历文件夹内的所有文件
    for filename in os.listdir(folder_path):
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            file_path = os.path.join(folder_path, filename)
            # 读取Excel文件并合并到combined_data
            if filename.endswith('.xlsx'):
                data = pd.read_excel(file_path,
                sheet_name=0,  #='表1'
                header=1, #表头为第二行
                converters={'单据号': str},
                usecols=['单据号','单据类型','日期','客户名称','药品名称','明细金额','成本','数量']) 
            else:
                data = pd.read_excel(file_path, engine='xlrd',
                sheet_name=0,  #='表1'
                header=1, #表头为第二行
                converters={'单据号': str},
                usecols=['单据号','单据类型','日期','客户名称','药品名称','明细金额','成本','数量'])   # 使用xlrd读取.xls文件
            combined_data = combined_data.append(data, ignore_index=True)
    combined_data['日期'] = pd.to_datetime(combined_data['日期'],format = '%Y-%m')
    combined_data['日期'] = combined_data['日期'].dt.strftime(("%Y-%m"))
    # 将合并后的数据写入输出文件
    combined_data.to_excel(output_file, index=False)
    print(f"合并完成，结果保存到 {output_file}")

# 设置文件夹路径和输出文件路径
folder_path = 'E:\Program Files (x86)\python\图表分析'  # 替换为你的文件夹路径
output_file = 'E:\Program Files (x86)\python\图表分析\输出\output.xlsx'       # 替换为输出文件的路径和名称
merge_excel_files(folder_path, output_file)


def bar_diagram(): #制作条形图
    df=combined_data.groupby(['客户名称'])['明细金额','成本'].sum().reset_index()
    [combined_data['明细金额'],combined_data['成本']]=[round(combined_data['明细金额'].astype("float64")/10000,2),round(combined_data['成本'].astype("float64")/10000,2)]
    df.sort_values(by="明细金额",inplace=True,ascending=False) #数据排序,按分数这列，直接修改数据，降序
    # 获取销售额前10的客户
    top_10_customers = df.head(10)
    # 计算其他客户的总销售额
    other_sales = df.iloc[10:]['明细金额'].sum()
    # 创建一个包含“其他”客户的DataFrame
    other_customer = pd.DataFrame({'客户名称': ['其他'], '明细金额': [other_sales]})
    # 将销售额前10的客户和“其他”客户合并
    df = pd.concat([top_10_customers, other_customer])
    #df.sort_values(by="明细金额",inplace=True,ascending=False) #数据排序,按分数这列，直接修改数据，降序
    # 解决坐标轴负号问题
    plt.rcParams['axes.unicode_minus'] = False
    # 画柱状图：x轴是金额，y轴是名称，颜色是红色
    plt.bar(x=0,bottom=df['客户名称'],height=0.5,width=df['明细金额'],label="单位:万元",orientation="horizontal",color="red",alpha=0.7)
    # lable的位置，左上解
    plt.legend(loc="upper right")
    # 显示图例
    # plt.legend()
    # 设置X与Y轴的标题
    plt.xlabel("明细金额",rotation=0)
    plt.ylabel("客户名称",rotation=0)
    # 刻度标签及文字旋转
    plt.xticks(df['明细金额'],rotation=45)
    #y轴的刻度范围
    #plt.xlim([-100, 5000])
    # 设置图表的标题、字号、粗体
    plt.title("销售明细",fontsize=16,fontweight='bold')
    # 把dataframe转换为list
    for x1, y1 in enumerate(df['明细金额']):
        plt.text(y1,x1, str(y1), ha='left',va='center', fontsize=12, color='black')
    # 紧凑型的布局
    plt.tight_layout()
    plt.savefig(r"E:\Program Files (x86)\python\图表分析\输出\柱状图.jpg",bbox_inches='tight', dpi=800)
    plt.show()
#bar_diagram()

def line_chart():
    # 1. 筛选数据
    # 仅保留客户名称、日期和明细金额这几列
    filtered_data = combined_data[['客户名称', '日期', '明细金额']]

    # 2. 制作透视表
    # 使用pivot_table函数制作透视表，列为客户名称，行为日期，值为明细金额的合计
    pivot_table = pd.pivot_table(filtered_data, values='明细金额', index='日期', columns='客户名称', aggfunc='sum', fill_value=0)

    # 3. 合并小客户为"其他"
    # 计算每个客户的总明细金额
    customer_total = pivot_table.sum(axis=0)

    # 找出总明细金额前3的客户
    top_3_customers = customer_total.nlargest(4).index

    # 创建"其他"列，将不在前3名的客户的合计金额汇总到"其他"列
    pivot_table['其他'] = pivot_table.drop(columns=top_3_customers).sum(axis=1)

    # 4. 删除不需要的客户列
    pivot_table.drop(columns=pivot_table.columns.difference(top_3_customers.union(['其他'])), inplace=True)

    # 打印透视表
    pivot_table.to_excel('输出\demo.xlsx')


    # 将日期列转换为日期时间对象
    pivot_table.index = pd.to_datetime(pivot_table.index)

    # 4. 汇总后的金额保留2位小数
    pivot_table = round(pivot_table/10000,2)


    # 绘制折线图
    #plt.figure(figsize=(12.8, 7.2))  # 设置图形大小
    for column in pivot_table.columns:
        plt.plot(pivot_table.index, pivot_table[column], label=column,alpha=0.5)



    for column in pivot_table.columns:
        for x,z in zip(pivot_table.index,pivot_table[column]):
            plt.text(x, z, str(z), ha='center', va='bottom', rotation=0)

    

    plt.tight_layout()
    plt.xlabel('日期')
    plt.ylabel('明细金额')
    plt.title('客户明细金额趋势图')
    plt.legend(loc='best')  # 添加图例
    plt.grid(True)  # 添加网格线
    plt.savefig(r"E:\Program Files (x86)\python\图表分析\输出\折线图.jpg",bbox_inches='tight', dpi=800)
    plt.show()
line_chart()