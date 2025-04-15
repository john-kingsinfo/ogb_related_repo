import json
import mysql.connector
import pandas as pd
import paras

ogbconn = mysql.connector.connect(
    # 連線主機名稱
    host=paras.ogb_host,
    # 登入帳號
    user=paras.ogb_user,
    # 登入密碼
    password=paras.ogb_pwd,
)
ogbc = ogbconn.cursor()

event_id = 306
field_list = []
orders = []
orders_columns = ['報名序號', '活動名稱', '報名日期', '報名狀態', '付款狀態', '付款金額']
options = {}
option_list = []
order_status = {0: '候補',
                1: '報名失敗',
                2: '報名成功',
                3: '待退刷',
                9: '已取消報名'}
pay_status = {0: '尚未付款',
              1: '已付款',
              2: '超過付款時間',
              3: '已取消',
              4: '待補差額'}

# load field base info

sql = f'''
SELECT 
    event_base_info, title
FROM
    OneGoBoLine.events
WHERE
    id = {event_id}
;
'''
ogbc.execute(sql)
r = ogbc.fetchone()
base_info_list = [v['name'] for v in json.loads(r[0]) if v['value']]
event_title = r[1]

sql = f'''
SELECT 
    name_en, name_ch
FROM
    OneGoBoLine.base_info_field
WHERE
    business_id = 1
ORDER BY id ASC
;
'''
ogbc.execute(sql)
for r in ogbc.fetchall():
    if r[0] in base_info_list:
        field_list.append(r[0])
        orders_columns.append(r[1])
# base_info_list = orders_columns[1:]
order_columns_base_len = len(orders_columns)


# load all options from event
sql = f'''
SELECT 
    eos.id, eos.name, eosi.id, eosi.name
FROM
    OneGoBoLine.event_options_settings AS eos
        LEFT JOIN
    OneGoBoLine.event_options_settings_inputs AS eosi ON eosi.events_options_settings_id = eos.id
WHERE
    events_id = {event_id}
    AND eos.valid = 1
ORDER BY eos.id, eosi.id ASC
;
'''
ogbc.execute(sql)
for r in ogbc.fetchall():
    if r[0] not in options:
        options[r[0]] = {'id': r[0], 'in_id': [], 'column': ['報名序號', '報名狀態'], 'orders': {}}
        options[r[0]]['name'] = r[1].replace('/', '')
        option_list.append(r[0])
        orders_columns.append(r[1])
    if r[2]:
        options[r[0]]['in_id'].append(r[2])
        options[r[0]]['column'].append(r[3])


# get orders from event id
sql = f'''
SELECT 
    eo.id, eo.created_at, eo.order_status, eo.pay_status, eo.pay_total, ea.base_info
FROM
    OneGoBoLine.event_order AS eo
        JOIN
    OneGoBoLine.event_attendees AS ea ON eo.id = ea.event_order_id
WHERE
    eo.events_id = {event_id}
ORDER BY eo.id ASC
;
'''
ogbc.execute(sql)
for r in ogbc.fetchall():
    order_info = [r[0], event_title, r[1], order_status.get(r[2], r[2]), pay_status.get(r[3], r[3]), r[4]]
    base_info = json.loads(r[5])
    for b in field_list:
        order_info.append(base_info.get(b, ''))
    # append options count with 0        
    order_info = order_info  + [0] * len(option_list)
    orders.append(order_info)
    

# for each event option create one sheet
for oid in options:
    in_id: list = options[oid]['in_id']
    ord_op_ind = option_list.index(oid) + order_columns_base_len
    columns_len = len(in_id) + 2
    sql = f'''
SELECT 
    eo.id,
    eo.order_status,
    eop.id,
    eoi.event_options_settings_inputs_id,
    eoi.value
FROM
    OneGoBoLine.event_order AS eo
        LEFT JOIN
    OneGoBoLine.event_options AS eop ON eop.event_order_id = eo.id
        LEFT JOIN
    OneGoBoLine.event_options_input AS eoi ON eoi.event_options_id = eop.id
WHERE
    eo.events_id = {event_id}
        AND eop.event_options_settings_id = {oid}
ORDER BY eo.id, eop.id ASC
;
    '''
    ogbc.execute(sql)
    p_opt_id = None
    opt_id = None
    for r in ogbc.fetchall():
        opt_id = r[2]
        if opt_id not in options[oid]['orders']:
            options[oid]['orders'][opt_id] = [''] * columns_len
            options[oid]['orders'][opt_id][0] = r[0]
            options[oid]['orders'][opt_id][1] = order_status.get(r[1], r[1])
        c_ind = in_id.index(r[3]) + 2
        options[oid]['orders'][opt_id][c_ind] = r[4]
        

        # check event option id
        if opt_id == p_opt_id:
            continue
        p_opt_id = opt_id
        # search order list with order id, add options count
        for i in range(len(orders)):
            if orders[i][0] == r[0]:
                orders[i][ord_op_ind] += 1
                break

all_df = [(pd.DataFrame(orders, columns=orders_columns), '報名資訊')]
for oid in options:
    all_df.append((pd.DataFrame([options[oid]['orders'][o] for o in options[oid]['orders'] ], columns=options[oid]['column']), options[oid]['name']))

with pd.ExcelWriter(f'{event_title}-報名資訊.xlsx') as writer:
    for df in all_df:
        df[0].to_excel(writer, sheet_name=df[1], index=False)
