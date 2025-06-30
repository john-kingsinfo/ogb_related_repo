import datetime
import json
import pandas as pd
import sqlite3

con = sqlite3.connect("db.sqlite3")
# con = sqlite3.connect("db_test.sqlite3")

#                   0         1        2          3           4          5          6          7           8          9
order_column = ['結帳編號', '時間', '客戶編號', '客戶姓名', '寵物編號', '寵物姓名', '服務人員', '項目編號', '項目類別', '項目分類',
                '項目名稱', '預設單價', '數量', '費用', 'OGB member', '電話號碼', 'OGB訂單單號', '指定單價', '單價折扣', '結帳折扣']
#                  10       11         12      13         14            15         16         17           18           19

payment_column = ['結帳編號', '結帳總金額', '付款方式', '付款金額', '還款', '時間', '備註']
outlay_column = ['編號', '時間', '名稱', '服務人員', '金額', '備註']

payment_type = {'Cash': '現金',
                'Remittance': '匯款',
                'CreditCard': '信用卡',
                'Arrears': '欠款',
                'Mobile': '多元支付'}

arrears = []
arrears_code = 'MI00002'
prepay = '已繳金額'

sql_datetime_formate = '%Y-%m-%d'
display_formate = '%Y/%m/%d %H:%M'
start_date = datetime.datetime(2025, 5, 24)
# dur = 14
# if dur > 0:
#     end_date = datetime.timedelta(days=dur) + start_date
# else:
#     end_date = datetime.datetime(2024, 4, 1)
end_date = datetime.datetime(2025, 5, 25)
# end_date = datetime.datetime.now()
start_date_str = start_date.strftime(sql_datetime_formate)
end_date_str = end_date.strftime(sql_datetime_formate)
all_df = []
all_order = []
all_payment = []
all_outlay = []

sql = f"select scd.single_checkout_id as checkout_id, scd.checkout_date, \
        sc.owner_id, sc.owner_name, sc.patient_id, sc.patient_name, user.name, \
        scd.item_code, ci.type, ci.category, ci.name, ci.value, scd.count, \
        scd.subtotal_price, sc.extra_info, sc.owner_phone, sc.identifier, scd.value, '-', '-', \
        scd.item_name \
        from single_checkout_detail as scd \
        join single_checkout as sc on sc.id = checkout_id \
        join user on user.id = scd.user_id \
        left join checkout_item as ci on ci.code = scd.item_code \
        where scd.checkout_date between '{start_date_str}' and '{end_date_str}'\
        and sc.status = 'Enable'\
        order by scd.id asc"

# print(sql)
ITEMNAM_IDX = 10
OGB_IDX = 14
TPV_IDX = 17
TPD_IDX = 18
PPD_IDX = 19
for r in con.execute(sql):
    lr = list(r)
    j = json.loads(lr[OGB_IDX])
    if j['line_userid']:
        lr[OGB_IDX] = 'Y'
    else:
        lr[OGB_IDX] = 'N'
    if lr[ITEMNAM_IDX] is None:
        lr[ITEMNAM_IDX] = lr[-1]
    lr.pop()
    if lr[11] and lr[11] * lr[12] > 0:
        lr[TPD_IDX] = int(lr[TPV_IDX] * 100 / lr[11])
        if lr[13] > 0:
            lr[PPD_IDX] = int(lr[13] * 100 / (lr[11] * lr[12]))
    if lr[7] == arrears_code:
        arrears.append(lr[0])
    if lr[10] == prepay:
        lr[8] = lr[8] + "_prepay"
    all_order.append(lr)

sql = f"select sc.id, sc.total_price, scpd.type, scpd.price, '', sc.checkout_date, sc.description \
        from single_checkout as sc \
        join single_checkout_payment_detail as scpd on scpd.single_checkout_id = sc.id \
        where scpd.price >= 0 and sc.checkout_date between '{start_date_str}' and '{end_date_str}' \
        and sc.status = 'Enable'\
        order by sc.id asc"

for r in con.execute(sql):
    lr = list(r)
    lr[2] = payment_type.get(lr[2], lr[2])
    if lr[0] in arrears:
        lr[4] = '還款'
    all_payment.append(lr)

sql = f"select o.id, o.outlay_date, o.title, user.name, o.total_price, o.description \
        from outlay_record as o \
        join user on user.id = o.user_id \
        where o.outlay_date between '{start_date_str}' and '{end_date_str}'"

for r in con.execute(sql):
    lr = list(r)
    all_outlay.append(lr)

df1 = pd.DataFrame(all_order, columns=order_column)
df2 = pd.DataFrame(all_payment, columns=payment_column)
df3 = pd.DataFrame(all_outlay, columns=outlay_column)
all_df.append((df1, f'結帳'))
all_df.append((df2, f'付款方式'))
all_df.append((df3, f'差額調整'))

with pd.ExcelWriter(f'{start_date_str}-{end_date_str}.xlsx') as writer:
    for df in all_df:
        df[0].to_excel(writer, sheet_name=df[1], index=False)
