"""独立构造的虚构测试账单。无输入账单，不读取个人文件。"""
import argparse
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
HEADERS=['交易时间','交易类型','交易对方','商品','收/支','金额(元)','支付方式','当前状态','交易单号','商户单号','备注']
def build_rows():
    rows=[]
    def add(day,hour,party,goods,amount,direction='支出',status='支付成功',kind='商户消费'):
        rows.append((datetime(2024,4,day,hour),kind,party,goods,direction,float(amount),'演示支付渠道',status))
    for day in range(1,31):
        add(day,8,'示例食堂A','虚构早餐',4+day%3)
        add(day,12,'示例食堂A','虚构午餐',9+day%4)
        add(day,18,'示例食堂B','虚构晚餐',11+day%3)
        if day%5==0:add(day,15,'示例奶茶店','虚构饮品',7)
        if day%6==0:add(day,21,'示例便利店','虚构加餐',6)
        if day%7==0:add(day,13,'示例快餐店','虚构外卖正餐',16)
        if day%4==0:add(day,9,'示例地铁','虚构交通',2)
    add(1,10,'演示资助者A','虚构资助',1600,'收入','已存入零钱','转账')
    add(16,10,'演示资助者B','虚构资助',800,'收入','已存入零钱','转账')
    add(4,14,'示例网购店','虚构全额退款订单',64,status='已全额退款')
    add(5,14,'示例网购店','虚构退款回流',64,'收入','已存入零钱','商户-退款')
    add(11,14,'示例网购店','虚构部分退款订单',46,status='已退款(¥13.00)')
    add(12,14,'示例网购店','虚构退款回流',13,'收入','已存入零钱','商户-退款')
    add(3,10,'示例校园卡服务','虚构预付充值',120)
    add(8,16,'示例羽毛球馆','虚构场地费',17,kind='扫二维码付款')
    add(19,10,'示例教材店','虚构教材',73)
    add(22,10,'示例洗衣服务','虚构洗衣',4)
    add(24,10,'演示同学A','转账备注:虚构分摊',21,kind='转账')
    add(26,10,'演示银行账户','虚构提现',50,'/','提现已到账','零钱提现')
    return sorted(rows,key=lambda r:r[0])
def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--out',default=str(Path(__file__).resolve().parents[1]/'examples/data/demo_bill.xlsx'));a=ap.parse_args()
    rows=build_rows();w=Workbook();s=w.active;s.title='虚构演示账单'
    preamble=['完全虚构的测试账单','微信昵称：[演示数据 虚构人物]',
              '起始时间：[2024-04-01 00:00:00] 截止时间：[2024-04-30 23:59:59]',
              '所有商户、人物、交易与日期均为测试设定，不代表任何人的消费经历。',
              '生成器不读取外部账单或个人文件。',f'共{len(rows)}笔记录']
    for line in preamble:s.append([line])
    while s.max_row<13:s.append(['演示说明'])
    s.append(HEADERS)
    for i,r in enumerate(rows,1):s.append(list(r)+[f'DEMO-TXN-{i:04d}',f'DEMO-MERCHANT-{i:04d}','虚构测试记录'])
    w.properties.creator='Demo Generator';w.properties.lastModifiedBy='Demo Generator'
    p=Path(a.out);p.parent.mkdir(parents=True,exist_ok=True);w.save(p)
    print(f'已生成完全虚构的演示账单，共 {len(rows)} 笔。')
if __name__=='__main__':main()
