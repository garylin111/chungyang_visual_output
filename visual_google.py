import streamlit as st
import pandas as pd
import pygsheets
import json
from google.oauth2 import service_account
from plotly.subplots import make_subplots
import plotly.graph_objects as go
import plotly.express as px

# 設置服務賬號信息
secret_dict = dict(st.secrets["secret_key"])
# 標準數内容
# secret_num_dict = dict(st.secrets["data"])


# Streamlit 页面布局和输入
def main():
    # 當標工更新，需要使用test2程式做更新轉換
    data_st = [
            ['1253馬達座', 10, 9.8361],
            ['1253馬達座', 20, 6.2937],
            ['151馬達座', 10, 5.8],
            ['151馬達座', 20, 6.2],
            ['18引擎連桿組', 30, 43],
            ['18引擎連桿組', 40, 85],
            ['3017-1驅動盤', 10, 6.9498],
            ['3017-1驅動盤', 20, 5.1429],
            ['3018驅動盤', 10, 12.6],
            ['338針棒盒(7針)-完成碼', 30, 4.2],
            ['338針棒盒(7針)-完成碼', 40, 11.2],
            ['338針棒盒(7針)-完成碼', 50, 3.6],
            ['771 支撐管', 10, 513],
            ['771 支撐管', 20, 60],
            ['771 支撐管', 40, 50],
            ['771 支撐管', 50, 12],
            ['7針基座(1297)-完成碼', 30, 4.8],
            ['7針基座(1297)-完成碼', 50, 3.2],
            ['7針基座(1297)-完成碼', 60, 21.6],
            ['7針基座(1297)-完成碼', 70, 2.7],
            ['HD6.1-1 上座', 10, 24.8],
            ['HD6.1-1 上座', 20, 34.2],
            ['HD6.1-1 下座', 10, 22],
            ['RC-38汽缸', 10, 12.6],
            ['RC-38汽缸', 30, 11],
            ['S-7排氣接管(烤漆)', 20, 27.2727],
            ['S-7排氣接管(烤漆)', 40, 24],
            ['S-7支架(烤漆)', 10, 24.6],
            ['S-7馬達蓋(烤漆)', 10, 5.9],
            ['S-7齒輪箱(烤漆)', 10, 7.5],
            ['S-7齒輪箱(烤漆)', 30, 3],
            ['SS10進檔手把', 10, 47.1],
            ['UX1.2右碟仔', 10, 7.0039],
            ['UX1.2右碟仔', 20, 12.6761],
            ['UX1.2右碟仔', 30, 18],
            ['VERTEX曲軸箱', 11, 10.46],
            ['傳動器', 20, 16.98],
            ['傳動器', 30, 80],
            ['切線板', 10, 40],
            ['切線板', 30, 17],
            ['前後擺臂', 10, 4.1],
            ['前後擺臂', 20, 6.9],
            ['前後擺臂-1', 10, 12.0805],
            ['前後擺臂-1', 20, 8.4706],
            ['前蓋(灰)', 10, 18.2],
            ['右轉向軸', 20, 5],
            ['天秤', 10, 33.3333],
            ['導板', 10, 27.4],
            ['導線板完成品', 20, 27.6923],
            ['左轉向軸', 20, 5.7052],
            ['左轉向軸', 40, 39],
            ['後擺臂', 10, 5.1429],
            ['後擺臂', 20, 27.907],
            ['截斷閥', 10, 5.3412],
            ['控制箱M版', 10, 6.8],
            ['擋料座', 10, 30],
            ['支撐架', 20, 33.3333],
            ['時序輪5469', 10, 33.5],
            ['時序輪5469', 20, 32],
            ['時序輪5469', 30, 13],
            ['時序輪5469', 40, 108],
            ['時序輪5469', 50, 700],
            ['時序輪5469', 60, 57],
            ['時序輪5469', 70, 314],
            ['時序輪5469', 80, 212],
            ['時序輪5469', 90, 360],
            ['時序輪5469', 100, 250],
            ['時序輪5469', 110, 220],
            ['曲軸(1205)', 10, 135],
            ['曲軸箱(P+R)', 10, 18.5],
            ['曲軸箱(P+R)', 50, 23.2],
            ['本體組', 20, 19.3548],
            ['汽缸', 10, 15.859],
            ['汽缸', 20, 29.5082],
            ['汽缸', 30, 40.4494],
            ['汽缸蓋(藍)', 30, 46.7532],
            ['汽缸蓋(藍)', 40, 57.1429],
            ['活塞', 20, 90],
            ['活塞', 30, 47.6],
            ['第一階行星臂', 10, 10],
            ['第二階行星臂', 10, 6.7039],
            ['葉蓋(灰)', 20, 60],
            ['轉向軸座', 10, 52],
            ['轉向軸座', 20, 7.1358],
            ['轉子6508', 10, 12],
            ['轉子6508', 20, 18],
            ['連栓(561)', 20, 10.2273],
            ['連桿', 20, 15.1261],
            ['連桿(561)', 20, 4],
            ['針基座(12針)', 20, 3.4253],
            ['針基座(12針)', 30, 3.5329],
            ['針基座(12針)', 40, 2.9221],
            ['針基座(15針)', 30, 3.5398],
            ['針基座(15針)', 40, 8.4507],
            ['針基座(15針)', 50, 2.4275],
            ['針板', 20, 170],
            ['針板', 70, 99],
            ['針棒曲軸(半成品)', 20, 9],
            ['鎖頭原色', 10, 16.4],
            ['鎖頭原色', 30, 115],
            ['飛輪(風扇)', 30, 97.3],
            ['馬達座371', 10, 7.045],
            ['馬達座371', 20, 4.4],
            ['馬達座371', 30, 45],
            ['齒輪箱蓋S-7', 30, 43.9],
            ['龍頭座', 20, 20.6897]
        ]
    data_st = pd.DataFrame(data_st, columns=['品名', '工序', '標準數'])
    st.title('生產計劃')

    start_time = st.date_input('訂單開始時間', key='start')
    end_time = st.date_input('訂單結束時間', key='end')
    google_url = st.text_input('輸入試算表鏈接',
                               'https://docs.google.com/spreadsheets/d/16i6O3ZWmbazC8_m14xj49A9U_sirJnRKrCzvtrLIvck/edit?gid=0#gid=0')

    if start_time and end_time and google_url:
        if st.button('提交', key='submit'):
            data = load_data(start_time, end_time, google_url)
            st.session_state['origin_data'] = data

    if 'origin_data' in st.session_state:
        with st.expander('目視化人員工時標工'):
            product_name(start_time, end_time, data_st)
            show_data()
        # 显示所有产品的产出
        with st.expander('顯示產品產出柱狀圖'):
            show_all_products_output(st.session_state['origin_data'])
        # if st.button('显示所有产品的产出'):
        #     show_all_products_output(st.session_state['origin_data'])

        # 统计时间段内各个产品的总产出
        with st.expander('各個產品總產出分佈'):
            show_total_outputs(st.session_state['origin_data'], start_time, end_time)
        # if st.button('统计时间段内各个产品的总产出'):
        #     show_total_outputs(st.session_state['origin_data'], start_time, end_time)


def load_data(start_time, end_time, url):
    SCOPES = ('https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive')
    service_account_info = json.dumps(secret_dict)

    credentials = service_account.Credentials.from_service_account_info(json.loads(service_account_info), scopes=SCOPES)

    gc = pygsheets.authorize(custom_credentials=credentials)

    sheet = gc.open_by_url(url)
    worksheet = sheet.sheet1
    data = worksheet.get_all_records()

    df = pd.DataFrame(data)
    df['製造日'] = pd.to_datetime(df['製造日'], errors='coerce')  # 将日期列转换为 datetime64[ns]

    # 转换输入的日期为 datetime64[ns]
    start_time = pd.to_datetime(start_time)
    end_time = pd.to_datetime(end_time)

    mask = (df['製造日'] >= start_time) & (df['製造日'] <= end_time)
    filtered_data = df.loc[mask]

    return filtered_data


def get_final_outputs(data):
    # 将工序列中的 NaN 值替换为一个较大的数，以确保它们被视为最后一道工序
    data['工序'] = data['工序'].fillna(data['工序'].max() + 1)

    # 找到每个产品的最后一道工序的产出作为最终产出
    final_outputs = data.groupby('品名')['工序'].max().reset_index()
    final_outputs = final_outputs.merge(data, on=['品名', '工序'], how='left')

    return final_outputs


# 显示数据和可视化
def get_chinese_weekday(date):
    weekdays = ['（一）', '（二）', '（三）', '（四）', '（五）', '（六）', '（日）']
    return weekdays[date.weekday()]


def show_data():
    origin_data = st.session_state['origin_data']

    st.header('早晚班產出目視化')

    # 添加篩選選項
    product_numbers = origin_data['料號'].unique()
    selected_product_number = st.selectbox('選擇料號', product_numbers)

    filtered_data_by_number = origin_data[origin_data['料號'] == selected_product_number]
    product_names = filtered_data_by_number['品名'].unique()
    selected_product_name = st.selectbox('選擇品名', product_names)

    filtered_data_by_name = filtered_data_by_number[filtered_data_by_number['品名'] == selected_product_name]
    process_names = filtered_data_by_name['工序'].unique()
    selected_process = st.selectbox('選擇工序', process_names)

    filtered_data_by_process = filtered_data_by_name[filtered_data_by_name['工序'] == selected_process]
    shifts = filtered_data_by_process['班別'].unique()
    selected_shift = st.selectbox('選擇班別', shifts)

    # 篩選數據
    filtered_data = filtered_data_by_process[filtered_data_by_process['班別'] == selected_shift]

    filtered_data['產出'] = pd.to_numeric(filtered_data['產出'], errors='coerce').fillna(0).astype(int)
    filtered_data['工時'] = pd.to_numeric(filtered_data['工時'], errors='coerce').fillna(0).astype(float)

    grouped_data = filtered_data.groupby(['製造日'], as_index=False)[['產出', '工時']].sum()
    grouped_data['標工'] = grouped_data['產出'] / grouped_data['工時']
    grouped_data['標工'].fillna(0, inplace=True)

    # 将日期格式化为 "月日 + 星期"
    grouped_data['製造日格式化'] = grouped_data['製造日'].dt.strftime('%m月%d日') + ' ' + grouped_data['製造日'].apply(
        get_chinese_weekday)

    # show一下dataframe内容
    st.write('展示數據', grouped_data)

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=grouped_data['製造日格式化'],
            y=grouped_data['產出'],
            name='總產出',
            text=grouped_data['產出'],
            textposition='inside',
            textfont=dict(color='black'),  # 更改字體顏色
            hovertext=grouped_data['製造日格式化'],
        ),
        secondary_y=False,
    )

    fig.add_trace(
        go.Scatter(
            x=grouped_data['製造日格式化'],
            y=grouped_data['標工'],
            name='標工',
            mode='lines+markers+text',
            line=dict(color='firebrick', width=2),
            marker=dict(size=6),
            text=grouped_data['標工'].round(2),
            textfont=dict(color='black'),  # 更改字體顏色
            textposition='top center',
        ),
        secondary_y=True,
    )

    fig.update_layout(
        title_text=f'{selected_product_number} - {selected_product_name} - {selected_process} - {selected_shift} 下的所有人員總產出和標工',
        barmode='group',
        xaxis_title='製造日',
        yaxis_title='總產出',
        yaxis2_title='標工',
        xaxis=dict(
            tickangle=-45  # 標籤角度
        )
    )

    st.plotly_chart(fig)


def product_name(start_time, end_time, standard):

    origin_data = st.session_state['origin_data']

    # 顯示標題
    st.header('產品工序人員排名')

    # 選擇產品、工序
    product_numbers = origin_data['料號'].unique()
    selected_product_number = st.selectbox('選擇料號', product_numbers, key='select_product_number')

    filtered_data = origin_data[origin_data['料號'] == selected_product_number]

    product_names = filtered_data['品名'].unique()
    selected_product_name = st.selectbox('選擇品名', product_names, key='select_product_name')

    process_names = filtered_data['工序'].unique()
    selected_process = st.selectbox('選擇工序', process_names, key='select_process')

    filtered_data = filtered_data[filtered_data['工序'] == selected_process]

    # 統計每個人員的工時和標工，包括班別和製造日
    grouped_data = filtered_data.groupby(['製造日', '班別', '姓名'], as_index=False).agg({
        '工時': 'sum',
        '產出': 'sum'
    })

    # 獲取標準工時值
    filtered_standard = standard[
        (standard['品名'] == selected_product_name) &
        (standard['工序'] == selected_process)
    ]

    if filtered_standard.empty:
        st.error(f"找不到標準工時值，請檢查品名和工序：{selected_product_name}, {selected_process}")
        return

    standard_num = filtered_standard['標準數'].values[0]

    grouped_data['標工'] = grouped_data['產出'] / grouped_data['工時']
    grouped_data['理論產出'] = grouped_data['工時'] * standard_num

    # 將理論產出保留到小數點後兩位
    grouped_data['標工'] = pd.to_numeric(grouped_data['標工'], errors='coerce').fillna(0).round(0)
    grouped_data['理論產出'] = pd.to_numeric(grouped_data['理論產出'], errors='coerce').fillna(0).round(0)

    # 將日期格式化為 "月日 + 星期"
    grouped_data['製造日格式化'] = grouped_data['製造日'].dt.strftime('%m月%d日') + ' ' + grouped_data['製造日'].apply(get_chinese_weekday)

    # 重新排序列
    grouped_data = grouped_data[['製造日', '製造日格式化', '姓名', '班別', '標工', '工時', '產出', '理論產出']]

    # 按製造日排序
    grouped_data = grouped_data.sort_values(by='製造日')

    # 顯示結果
    st.header(f'{start_time}~{end_time}')

    st.header(f'{selected_product_name} - 工序 {selected_process}')
    st.write('標準數：', standard_num)
    st.dataframe(grouped_data)

    # 使用 Plotly 可視化
    fig = go.Figure()

    # 定義班別顏色
    shift_colors = {
        '早班': 'blue',
        '晚班': 'red'
    }

    for shift, color in shift_colors.items():
        shift_data = grouped_data[grouped_data['班別'] == shift]
        fig.add_trace(go.Bar(
            x=shift_data['製造日格式化'],
            y=shift_data['產出'],
            name=f'產出 - {shift}',
            marker_color=color,
            opacity=0.7,
            text=shift_data.apply(lambda row: f"{row['姓名']}<br>產出: {row['產出']}", axis=1),
            hovertemplate='<b>%{x}</b><br>%{text}<extra></extra>',
        ))

        fig.add_trace(go.Bar(
            x=shift_data['製造日格式化'],
            y=shift_data['理論產出'],
            name=f'理論產出 - {shift}',
            marker_color=color,
            opacity=0.3,
            text=shift_data.apply(lambda row: f"{row['姓名']}<br>理論產出: {row['理論產出']}", axis=1),
            hovertemplate='<b>%{x}</b><br>%{text}<extra></extra>',
        ))

    fig.update_layout(
        title=f'{selected_product_name} - 工序 {selected_process}',
        xaxis_title='製造日',
        yaxis_title='產出與理論產出',
        barmode='group',  # 分組顯示
        legend_title_text='指標'
    )

    st.plotly_chart(fig)

def show_all_products_output(data):
    st.header('所有產品的產出')

    grouped_data = data.groupby(['製造日', '品名'], as_index=False)['產出'].sum()

    # 将日期格式化为 "月日 + 星期"
    grouped_data['製造日格式化'] = grouped_data['製造日'].dt.strftime('%m月%d日') + ' ' + grouped_data['製造日'].apply(get_chinese_weekday)

    fig = go.Figure()

    for product in grouped_data['品名'].unique():
        product_data = grouped_data[grouped_data['品名'] == product]
        fig.add_trace(
            go.Bar(
                x=product_data['製造日格式化'],
                y=product_data['產出'],
                name=f'{product} 產出',
                text=product_data['產出'],
                textposition='outside',
                hovertext=product_data['品名'],
                legendgroup=product,
            )
        )

    fig.update_layout(
        title_text='所有產品的產出隨時間變化',
        barmode='group',
        xaxis_title='製造日',
        yaxis_title='總產出',
        xaxis=dict(
            tickangle=-45  # 標籤角度
        )
    )

    st.plotly_chart(fig)



def show_total_outputs(data, start_time, end_time):
    st.header(f'{start_time}-{end_time} 各個產品的總產出')
    final_outputs = get_final_outputs(data) # 使用最後的工序作爲產品產出
    grouped_data = final_outputs.groupby('品名')['產出'].sum().reset_index()

    col1, col2 = st.columns(2)

    with col1:
        st.write(grouped_data)

    with col2:
        fig = px.pie(grouped_data, values='產出', names='品名', title='各產品的產出占比')
        st.plotly_chart(fig)


if __name__ == '__main__':
    main()
