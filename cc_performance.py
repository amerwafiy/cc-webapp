import streamlit as st
import pandas as pd
import datetime as dt
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from PIL import Image
import matplotlib.pyplot as plt
from colour import Color

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

image = Image.open('invoke_logo.jpg')
st.image(image)

st.title('Call Centre Performance Web App')


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Data')
    workbook = writer.book
    worksheet = writer.sheets['Data']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    return processed_data


def generate_result(df):
    df = df.sort_values('Call Start DT').reset_index(drop=True)
    df = df[df['Dial Leg'] == 'agent']
    df = df[['Agent Username', 'Call Start DT', 'Call Dur Full']]
    df['Call Start DT'] = [x[:10] for x in df['Call Start DT']]
    df['Call Start DT'] = pd.to_datetime(
        df['Call Start DT'], format='%Y/%m/%d').apply(lambda x: dt.datetime.strftime(x, '%d/%m/%Y'))
    agents = list(df['Agent Username'].unique())
    dates = list(df['Call Start DT'].unique())
    result = pd.DataFrame({'Agent ID': agents})
    for d in dates:
        sub_date = df[df['Call Start DT'] == d]
        temp1 = []
        temp2 = []
        temp3 = []
        temp4 = []
        for agent in list(result['Agent ID']):
            agent_df = sub_date[sub_date['Agent Username'] == agent]
            total_calls = len(agent_df)
            cr = agent_df[agent_df['Call Dur Full'] > 59]

            total_duration = round(sum(list(agent_df['Call Dur Full'])) / total_calls)
            total_duration_cr = round(sum(list(cr['Call Dur Full'])) / len(cr))
            temp1.append(total_calls)
            temp2.append(total_duration)
            temp3.append(len(cr))
            temp4.append(total_duration_cr)

        result[d + ': # Calls Attempted'] = temp1
        result[d + ': Avg Call Duration(s)'] = temp2
        result[d + ': # CR'] = temp3
        result[d + ': Avg CR Duration(s)'] = temp4
    result = result.sort_values('Agent ID').reset_index(drop=True)
    return result


def generate_chart(x, y, date):
    green = Color("green")
    colors = list(green.range_to(Color("red"), len(x)))
    colors = [color.rgb for color in colors]
    fig, ax = plt.subplots()
    ax.barh(x, y,  align='center', color=colors)
    ax.invert_yaxis()  # labels read top-to-bottom
    ax.set_xlabel('Number of Calls')
    ax.set_title('Number of Calls Attempted By CC Agent ' + '('+date+')')
    st.pyplot(fig)


uploaded_file = st.file_uploader("Upload a file (csv/xlsx)")


if uploaded_file is not None:
    st.subheader('Number of calls attempted')
    filename = uploaded_file.name

    if filename[-3:] == 'csv':
        # Can be used wherever a "file-like" object is accepted:
        df = pd.read_csv(uploaded_file, header=5)

        with st.spinner('Wait for it...'):
            result = generate_result(df)
            st.table(result)

        df_xlsx = to_excel(result)
        st.download_button(label='📥 Download Result',
                           data=df_xlsx,
                           file_name='cc_report_' + result.columns[-1] + '.xlsx')
        if st.checkbox('Visualize Result'):
            for d in list(result.columns[1:]):
                result_ordered = result.sort_values(d, ascending=False)
                with st.spinner('Wait for it...'):
                    generate_chart(result_ordered['Agent ID'], result_ordered[d], d)
    elif filename[-4:] == 'xlsx':
        with st.spinner('Wait for it...'):
            df = pd.read_excel(uploaded_file, header=5)

        result = generate_result(df)
        st.table(result)

        df_xlsx = to_excel(result)
        st.download_button(label='📥 Download Result',
                           data=df_xlsx,
                           file_name='cc_report_' + result.columns[-1] + '.xlsx')
        if st.checkbox('Visualize Result'):
            for d in list(result.columns[1:]):
                result_ordered = result.sort_values(d, ascending=False)
                with st.spinner('Wait for it...'):
                    generate_chart(result_ordered['Agent ID'], result_ordered[d], d)
    else:
        st.write('Please upload a csv or xlsx file only')
