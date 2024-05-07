import os
import re
import pandas as pd
#from rich import print
import xmltodict
from dataclasses import dataclass,field

DF_HEADER_DICT = {'order_id':'订单编号',
                  'order_time':'下单日期',
                  'order_source':'来源平台',
                  'real_income':'商家实收金额',
                  'postage_fee':'总配送费',
                  'meituan_fee':'美团推广成本',
                  'street_promoter':'地推成本',
                  'single_profit':'客单利润',
                  'commission':'总部抽成',
                  }

@dataclass(order=True)
class OrderInfo(object):
    order_id        : str = ''
    order_time      : str = ''
    order_source    : str = ''
    real_income     : float = 0.0
    postage_fee     : float = 0.0
    meituan_fee     : float = 2.5
    street_promoter : float = 0.0
    single_profit   : float = 0.0
    commission      : float = 0.0

    def __post_init__(self):
        if(self.postage_fee in [0.0, 0, '', '0']):
            self.postage_fee = 0
            self.street_promoter = 15
        self.calc_single_profit(self.postage_fee)
        self.commission = self.real_income/0.9 * 0.1

    def calc_single_profit(self, is_real):
        if(is_real!=0):
            self.single_profit = self.real_income - (self.postage_fee+self.meituan_fee+self.street_promoter)
        else:
            self.single_profit = 0 - (self.postage_fee+self.meituan_fee+self.street_promoter)

@dataclass(order=True)
class OrderTable(object):
    real_orders     : int = 0
    fake_orders     : int = 0
    total_income    : float = 0.0
    average_income  : float = 0.0
    reback_income   : float = 0.0
    table           : list[OrderInfo] = field(default_factory=list)
    real_table      : list[OrderInfo] = field(default_factory=list)
    promoter_table  : list[OrderInfo] = field(default_factory=list)

    def __post_init__(self):
        self.split_real_orders()
        self.real_orders = self.get_real_order()
        self.fake_orders = len(self.table) - self.real_orders
        self.total_income = self.get_totoal_income()
        self.average_income = self.total_income/self.real_orders
        self.reback_income = self.get_reback_income()

    def __repr__(self):
        return f'实际单量：{self.real_orders}, \n实际总收入：{self.total_income}, \n平均单利润:{self.average_income}'

    def split_real_orders(self):
        for orderinfo in self.table:
            if ( not orderinfo.postage_fee == 0):
                self.real_table.append(orderinfo)
            else:
                self.promoter_table.append(orderinfo)

    def get_real_order(self):
        count = len(self.real_table)
        return count

    def get_totoal_income(self):
        total_income = 0.0
        for orderinfo in self.real_table:
            total_income = total_income + orderinfo.real_income
        for orderinfo in self.promoter_table:
            total_income = total_income - (orderinfo.meituan_fee+orderinfo.street_promoter)
        return total_income

    def get_reback_income(self):
        reback_income = 0.0
        for order in self.promoter_table:
            reback_income = reback_income + order.commission
        return reback_income


def find_excel_file():
    file_path = './'  # 指定文件夹路径
    file_ext = ['.xls']  # 文件扩展名列表
    file_list = []  # 存储文件路径的列表

    for path in os.listdir(file_path):
        path_list = os.path.join(file_path, path)  # 连接当前目录及文件或文件夹名称
        if os.path.isfile(path_list):  # 判断当前路径是否是文件
            if os.path.splitext(path_list)[1] in file_ext:  # 判断文件的扩展名是否是.xls、.xlsx
                file_list.append(path_list)  # 将文件路径添加到列表中

    print(f"目录下共有 {len(file_list)} 个 xls 文件。")
    print(file_list)
    return file_list

def xml2xlsx(xml):
    rows = []
    xml_data = open(xml, 'r', encoding='utf-8').read()
    parsed_data = xmltodict.parse(xml_data)
    rows = parsed_data['Workbook']['Worksheet']['Table']['Row']

    row_list = []
    for row in rows:
        row_item_list = []
        for item in row['Cell']:
            if('#text' not in item['Data'].keys()):
                row_item_list.append(0)
            else:
                row_item_list.append(item['Data']['#text'])
        #print(row_item_list)
        row_list.append(row_item_list)

    pd.set_option('display.max_rows', 500)
    df = pd.DataFrame(data=row_list[1:len(row_list)], columns=row_list[0])
    return df


def process_cake_table(df):
    orderinfo_list = []
    for index, row in df.iterrows():
        orderinfo = OrderInfo(
            order_id=row.loc[DF_HEADER_DICT['order_id']],
            order_time=row.loc[DF_HEADER_DICT['order_time']],
            order_source=row.loc[DF_HEADER_DICT['order_source']],
            real_income=float(row.loc[DF_HEADER_DICT['real_income']]),
            postage_fee=float(row.loc[DF_HEADER_DICT['postage_fee']]),
        )
        orderinfo_list.append(orderinfo)
    #print(orderinfo_list)

    ordertable = OrderTable(
        table = orderinfo_list
    )
    print(ordertable)
    return ordertable

def save_to_excel(ori_file_name, ordertable):
    ret = re.search("(./订单数据.*)-全部门店", ori_file_name)
    xlsx_name = ret.group(1)+'分析'+'.xlsx'
    #print(xlsx_name)
    writer = pd.ExcelWriter(xlsx_name)

    tmp_order = ordertable.table[0]
    header_list = [DF_HEADER_DICT[key] for key in tmp_order.__dict__.keys()]
    #print(header_list)

    real_table_list=[]
    for order in ordertable.real_table:
        real_table_list.append(order.__dict__.values())
    df = pd.DataFrame(data=real_table_list, columns=header_list)
    df.to_excel(writer, sheet_name='真实单', index=False)

    promoter_table_list=[]
    for order in ordertable.promoter_table:
        promoter_table_list.append(order.__dict__.values())
    df = pd.DataFrame(data=promoter_table_list, columns=header_list)
    df.to_excel(writer, sheet_name='刷单', index=False)

    sum_list = [[ordertable.real_orders, ordertable.fake_orders, ordertable.total_income, ordertable.average_income, ordertable.reback_income]]
    sum_header = ['真实单量', '刷单量', '减除配送费刷单费推广费的总利润', '每单平均利润', '总部需返点']
    df = pd.DataFrame(data=sum_list, columns=sum_header)
    df.to_excel(writer, sheet_name='统计', index=False)

    writer.close()
    return xlsx_name





if __name__ == '__main__':
    xml_list = find_excel_file()
    for f_xml in xml_list:
        row_list = xml2xlsx(f_xml)
        ordertable = process_cake_table(row_list)
        save_to_excel(f_xml, ordertable)
