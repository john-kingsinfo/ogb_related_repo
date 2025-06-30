import datetime
import mysql.connector
import pandas as pd
import paras

ezconn = mysql.connector.connect(
    # 連線主機名稱
    host=paras.ez_host,
    # 登入帳號
    user=paras.ez_user,
    # 登入密碼
    password=paras.ez_pwd,
)
ezc = ezconn.cursor()


# ezPretty 每日報表匯出測試
paymethod = {1: '店內信用卡刷卡',
             2: '現金',
             4: '課程或票券/OGB Ticket',
             5: '匯款',
             7: '定金',
             9: '折扣差額(非現金)/OGB會員折扣',
             10: 'One GoBo Points',
             11: '其他'
            }

store_id = {1: '天母店',
            2: '沐沐店',
            3: '吉林店',
            # 8: '測試店家'
           }

service_status = {2: '已預約',
                  3: '已完成'}

gender = {0: '女',
          1: '男'}

designers = {}

order_column = ['訂單編號', '開始時間', '結束時間', '客戶', '性別', '服務人員', '狀態', '服務', '分類', '項目', '單價', '數量', 
                '費用', '備註', 'OGB member', '寵物名', '種類', '品種', 'Line ID', '訂單建立時間', '結帳時間', '折扣後費用', '最低單價', '收費比例']
payment_column = ['訂單編號', '付款方式', '金額', '時間', '發票號碼']


sql_datetime_formate = '%Y-%m-%d'
display_formate = '%Y/%m/%d %H:%M'
start_date = datetime.datetime(2025, 1, 1)
dur = 11
if dur > 0:
    end_date = datetime.timedelta(days=dur) + start_date
else:
    end_date = datetime.datetime(2024, 4, 1)
end_date = datetime.datetime(2025, 6, 1)
# end_date = datetime.datetime.now()
start_date_str = start_date.strftime(sql_datetime_formate)
end_date_str = end_date.strftime(sql_datetime_formate)
all_df = []

# load designer name
sql = 'select id, name from Ezpretty.designers'
ezc.execute(sql)
result = ezc.fetchall()

for id, name in result:
    designers[id] = name

for the_store_id in store_id.keys():
    all_order = []
    all_payment = []
    payment_info = {}

    sql = f'select dbp.designer_booking_id, if(dbp.paymethod = 1, if(spc.pay_channel_id = 1, 1, 11), dbp.paymethod),\
        dbp.fee, db.start_time, db.invoice_number \
        from Ezpretty.designer_booking_payments as dbp\
        join Ezpretty.designer_bookings as db on dbp.designer_booking_id = db.id\
        LEFT JOIN Ezpretty.store_pay_channels AS spc ON dbp.store_pay_channel_id = spc.id\
        where db.start_time >= "{start_date_str}" and db.start_time < "{end_date_str}"\
        and db.status in (2, 3) and db.store_id = {the_store_id} and db.fee = db.total_fee'

    ezc.execute(sql)
    result = ezc.fetchall()
    for r in result:
        lr = list(r)
        if lr[0] not in payment_info:
            payment_info[lr[0]] = {'pay':0, 'discount': 0}
        lr[2] = int(lr[2])
        if lr[1] != 9:
            payment_info[lr[0]]['pay'] += lr[2]
        else:
            payment_info[lr[0]]['discount'] += lr[2]
        lr[1] = paymethod.get(lr[1], '未知')
        all_payment.append(lr)
    
    for k in payment_info.keys():
        if payment_info[k]['discount'] > 0:
            payment_info[k]['discount_percent'] = float(payment_info[k]['pay']) / (payment_info[k]['pay'] + payment_info[k]['discount'])

    all_payment = sorted(all_payment, key=lambda x: (x[3], x[0]))
    df2 = pd.DataFrame(all_payment, columns=payment_column)

    # query services order info
    sql = f'select db.id, db.start_time, dbs.elapsed_time, db.user_name, if(sc.user_gender = 0, "女", "男"), dbs.operation_designer_id\
        , db.status, "技術", dbs.name, dbs.item_name, dbs.price, dbs.quantity, dbs.total_fee\
        , db.afterservcie_comment, if(length(db.goboLineUserId) > 0, "Y", "N"), pro.name, if(prob.property_kind_id=1, "狗", "貓"), proc.name\
        , sc.line_user_id, db.created, db.pay_time, dbs.total_fee, ssit.mini_price, 1\
        FROM Ezpretty.designer_bookings as db\
        join Ezpretty.designers as des on db.designer_id = des.id\
        join Ezpretty.designer_booking_services as dbs on db.id = dbs.designer_booking_id\
        left join Ezpretty.property as pro on pro.id = dbs.store_customer_property_id\
        left join Ezpretty.property_category as proc on proc.id = pro.category_id\
        left join Ezpretty.property_bodytype as prob on prob.id = pro.bodytype_id\
        join Ezpretty.store_customers as sc on db.store_customer_id = sc.id\
        left join Ezpretty.store_service_item_templates as ssit on ssit.id = dbs.store_service_item_template_id\
        where db.start_time >= "{start_date_str}" and db.start_time <= "{end_date_str}"\
        and db.status in (2, 3) and db.store_id = {the_store_id} and db.fee = db.total_fee\
        order by db.id, dbs.id;'

    ezc.execute(sql)
    result = ezc.fetchall()
    MIN_PRICE_IDX = 20
    PAY_RATIO_IDX = 21
    oid = 0
    ST_IDX = 1
    ET_IDX = 2
    STAFF_IDX = 5
    S_STAT_IDX = 6
    PRICE_IDX = 10
    AMOUNT_IDX = 11
    T_PRICE_IDX = 12
    C_TIME_IDX = 19
    P_TIME_IDX = 20
    DIS_PRICE_IDX = 21
    MIN_PRICE_IDX = 22
    PAY_RATIO_IDX = 23
    for r in result:
        lr = list(r)
        if oid != lr[0]:
            oid = lr[0]
            st = lr[ST_IDX]
        lr[ST_IDX] = datetime.datetime.strftime(lr[ST_IDX], display_formate)
        st += datetime.timedelta(minutes=lr[ET_IDX] * int(lr[AMOUNT_IDX]))
        lr[ET_IDX] = datetime.datetime.strftime(st, display_formate)
        lr[STAFF_IDX] = designers[lr[STAFF_IDX]]
        lr[S_STAT_IDX] = service_status[lr[S_STAT_IDX]]
        lr[PRICE_IDX] = int(lr[PRICE_IDX])
        lr[AMOUNT_IDX] = int(lr[AMOUNT_IDX])
        lr[T_PRICE_IDX] = int(lr[T_PRICE_IDX])
        lr[C_TIME_IDX] = datetime.datetime.strftime(lr[C_TIME_IDX], display_formate)
        lr[P_TIME_IDX] = datetime.datetime.strftime(lr[P_TIME_IDX], display_formate) if lr[P_TIME_IDX] else ''
        if lr[0] in payment_info and payment_info[lr[0]]['discount'] > 0:
            lr[DIS_PRICE_IDX] = int(lr[DIS_PRICE_IDX] * payment_info[lr[0]]['discount_percent'])
        else:
            lr[DIS_PRICE_IDX] = int(lr[DIS_PRICE_IDX])
        lr[MIN_PRICE_IDX] = int(lr[MIN_PRICE_IDX]) if lr[MIN_PRICE_IDX] else 0
        lr[PAY_RATIO_IDX] = lr[DIS_PRICE_IDX] / (lr[AMOUNT_IDX] * lr[MIN_PRICE_IDX]) if lr[AMOUNT_IDX] * lr[MIN_PRICE_IDX] > 0 else 1
        all_order.append(lr)

    # query goods order info
    sql = f'select db.id, db.start_time, db.start_time, db.user_name, if(sc.user_gender = 0, "女", "男"), des.name\
        , db.status, "商品", dbg.goods_category_name, dbg.goods_name, dbg.sold_price, dbg.quantity, dbg.goods_total_fee\
        , db.afterservcie_comment, if(length(db.goboLineUserId) > 0, "Y", "N"), "", "", "", sc.line_user_id, db.created, db.pay_time\
        , dbg.goods_total_fee, goods.goods_cost, 1\
        FROM Ezpretty.designer_bookings as db\
        join Ezpretty.designers as des on db.designer_id = des.id\
        join Ezpretty.store_customers as sc on db.store_customer_id = sc.id\
        join Ezpretty.designer_booking_goods as dbg on db.id = dbg.designer_booking_id\
        left join Ezpretty.goods as goods on goods.id = dbg.goods_id\
        where db.start_time >= "{start_date_str}" and db.start_time <= "{end_date_str}"\
        and db.status in (2, 3) and db.store_id = {the_store_id} and db.fee = db.total_fee\
        order by db.id;'
    # print(sql)
    ezc.execute(sql)
    result = ezc.fetchall()

    for r in result:
        lr = list(r)
        lr[ST_IDX] = datetime.datetime.strftime(lr[ST_IDX], display_formate)
        lr[ET_IDX] = datetime.datetime.strftime(lr[ET_IDX], display_formate)
        lr[S_STAT_IDX] = service_status[lr[S_STAT_IDX]]
        lr[PRICE_IDX] = int(lr[PRICE_IDX])
        lr[AMOUNT_IDX] = int(lr[AMOUNT_IDX])
        lr[T_PRICE_IDX] = int(lr[T_PRICE_IDX])
        lr[C_TIME_IDX] = datetime.datetime.strftime(lr[C_TIME_IDX], display_formate)
        lr[P_TIME_IDX] = datetime.datetime.strftime(lr[P_TIME_IDX], display_formate) if lr[P_TIME_IDX] else ''
        lr[MIN_PRICE_IDX] = int(lr[MIN_PRICE_IDX]) if lr[MIN_PRICE_IDX] else 0
        lr[PAY_RATIO_IDX] = lr[DIS_PRICE_IDX] / (lr[AMOUNT_IDX] * lr[MIN_PRICE_IDX]) if lr[AMOUNT_IDX] * lr[MIN_PRICE_IDX] > 0 else 1
        all_order.append(lr)

    all_order = sorted(all_order, key=lambda x: (x[1], x[0]))
    df1 = pd.DataFrame(all_order, columns=order_column)

    all_df.append((df1, f'{store_id[the_store_id]}_訂單'))
    all_df.append((df2, f'{store_id[the_store_id]}_付款方式'))

with pd.ExcelWriter(f'ez_{start_date_str}-{end_date_str}.xlsx') as writer:
    for df in all_df:
        df[0].to_excel(writer, sheet_name=df[1], index=False)
