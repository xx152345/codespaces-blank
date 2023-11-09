import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import time
from matplotlib.ticker import FuncFormatter, ScalarFormatter
import matplotlib as mpl
from matplotlib.ticker import MaxNLocator

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
    [combined_data['明细金额'],combined_data['成本']]=[round(combined_data['明细金额'].astype("float64"),2),round(combined_data['成本'].astype("float64"),2)]
    # 将合并后的数据写入输出文件
    combined_data.to_excel(output_file, index=False)
    print(f"合并完成，结果保存到 {output_file}")

# 设置文件夹路径和输出文件路径
folder_path = 'E:\Program Files (x86)\python\图表分析'  # 替换为你的文件夹路径
output_file = 'E:\Program Files (x86)\python\图表分析\输出\output.xlsx'       # 替换为输出文件的路径和名称
merge_excel_files(folder_path, output_file)
customers = combined_data['客户名称'].unique()
def groupbys(combined_data):

    df=combined_data.groupby(['客户名称','日期'])['明细金额','成本'].sum().reset_index()
    #[combined_data['明细金额'],combined_data['成本']]=[round(combined_data['明细金额'].astype("float64")/10000,2),round(combined_data['成本'].astype("float64")/10000,2)]
    [combined_data['明细金额'],combined_data['成本']]=[round(combined_data['明细金额'].astype("float64"),2),round(combined_data['成本'].astype("float64"),2)]
    df.sort_values(by="明细金额",inplace=True,ascending=False) #数据排序,按分数这列，直接修改数据，降序
    return df

def bar_diagram(customer_data): #制作条形图
    # 解决坐标轴负号问题
    plt.rcParams['axes.unicode_minus'] = False
    # 画柱状图：x轴是金额，y轴是名称，颜色是红色
    plt.bar(x=0,bottom=customer_data['日期'],height=0.5,width=customer_data['明细金额'],label="单位:元",orientation="horizontal",color="red",alpha=0.7)
    #plt.bar(customer_data['日期'],customer_data['明细金额'],label="单位:元",color="red",alpha=0.7)

    # lable的位置，左上解
    plt.legend(loc="upper right")
    # 显示图例
    # plt.legend()
    # 设置X与Y轴的标题
    plt.xlabel('明细金额',rotation=0)
    plt.ylabel('日期',rotation=0)
    # 刻度标签及文字旋转
    plt.xticks(customer_data['明细金额'],rotation=30)
    #y轴的刻度范围
    #plt.xlim([-100, 5000])
    # 设置图表的标题、字号、粗体
    plt.title(customer+"销售明细",fontsize=16,fontweight='bold')
    # 把dataframe转换为list
    for x1, y1 in enumerate(customer_data['明细金额']):
        plt.text(y1,x1, format(round(y1,2),','), ha='left',va='top', fontsize=12, color='black')
    # 紧凑型的布局

    plt.tight_layout()
    plt.savefig(r"输出/"+customer+"柱状图.jpg",bbox_inches='tight', dpi=800)
    #plt.show()
    plt.close()
#bar_diagram()



for customer in customers:
    customer_data = combined_data[combined_data['客户名称'] == customer]
    customer_data=customer_data.groupby(['客户名称','日期'])['明细金额','成本'].sum().reset_index()
    combined_data['明细金额']=round(combined_data['明细金额'].astype("float64"),2)
    combined_data['成本']=round(combined_data['成本'].astype("float64"),2)
    customer_data.sort_values(by="日期",inplace=True,ascending=False) #数据排序,按分数这列，直接修改数据，降序
    print(customer_data)
    bar_diagram(customer_data)