# -*- coding: utf8 -*-
import warnings
warnings.filterwarnings('ignore')

import os
import shutil
from glob import glob
import pandas as pd
from tqdm import tqdm
from fpdf import FPDF
import time
# from pdf2image import convert_from_path


def cut_df_row(inp_df, inter_col, cut_val):
    # find latest row index
    last_row_index = -1
    for i, n in enumerate(inp_df[inter_col]):
        if int(n) == cut_val:
            last_row_index = i
            break

    return last_row_index

def preprocess_table(data):
    dfs = pd.read_excel(data, sheet_name=None) # sheet dict
    sheets = list(dfs.keys()) # sheet names
    # find source sheet
    src_sheet_name = [x for x in sheets if len(x.split('-')) == 2][0] # date

    # get src sheet, delete total sheet
    df = dfs[src_sheet_name]
    df.columns = list(df.iloc[0, :]) # first row to column
    df = df.loc[1:, :'총지급액(a-b+c)'].reset_index(drop=True) # slice column
    df = df.loc[:, ~df.columns.duplicated()].copy() # drop duplicate column
    df.columns = [x.strip().replace('\n', '') for x in df.columns]
    df = df.fillna(0) # nan to 0
    src_last_row_index = cut_df_row(inp_df=df, inter_col='No', cut_val=0)  # cut junk rows
    df = df.iloc[:src_last_row_index].reset_index(drop=True)
    df = df.fillna(0) # nan to 0

    # find detail sheet
    detail_sheet_name = [x for x in sheets if x=='상세내역'][0]
    df_detail = dfs[detail_sheet_name]
    user_no = sorted(list(df['No'].unique()))
    df_detail = df_detail[df_detail['No'].isin(user_no)]
    df_detail = df_detail.loc[:, :'공제1']
    df_detail = df_detail.fillna(0) # nan to 0
    df_detail = df_detail.reset_index(drop=True)

    return df, df_detail, src_sheet_name

def make_user_adjustment(user_no, save_dir):
    udf = df[df['No'] == user_no].reset_index(drop=True) # source talbe
    ddf = df_det[df_det['No'] == user_no].reset_index(drop=True) # content table

    oil_pay = int(round(udf.iloc[0, -2], 1)) # 유류대
    uinfo = udf[info_cols].values[0] # 개인기본정보

    # 지급 / 공제 df
    pay_udf = udf[payment_cols].T.reset_index() 
    ded_udf = udf[deduction_cols].T.reset_index()
    pay_udf.columns = ['P_Key', 'P_Value']
    ded_udf.columns = ['D_Key', 'D_Value']

    # 상세내역 추가
    pay_udf['P_Content'] = ddf[payment_cols].values[0]
    ded_udf['D_Content'] = ddf[deduction_cols].values[0]

    # 지급, 공제 테이블 병합
    merge_df = pd.concat([pay_udf, ded_udf], axis=1).fillna(0)
    merge_df['P_Value'] = merge_df['P_Value'].astype('int')
    merge_df['D_Value'] = merge_df['D_Value'].astype('int')

    # 지급 총액 / 공제 총액 / 실지급액 / 총지급액 계산
    pay_sum = int(merge_df['P_Value'].sum())
    ded_sum = int(merge_df['D_Value'].sum())
    actual_pay =  int(pay_sum-ded_sum)
    total_pay = int(actual_pay + oil_pay)

    # 급액 세자리 ',' 표기
    merge_df['P_Value'] = [format(x, ',') for x in merge_df['P_Value']]
    merge_df['D_Value'] = [format(x, ',') for x in merge_df['D_Value']]

    # 표출값 변환
    values = merge_df.replace(replace_dict).values

    ##################### Make pdf file
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_font('NanumGothic', '', './fonts/NanumGothic-Regular.ttf', uni=True)
    pdf.add_font('NanumGothicBold', '', './fonts/NanumGothic-Bold.ttf', uni=True)
    pdf.add_page()

    ###### 운송비 정산내역
    pdf.set_font('NanumGothicBold', '', 20)
    line_height = pdf.font_size * 1.0
    col_width = pdf.epw / 1
    pdf.multi_cell(col_width, line_height, '운송비 정산내역', border=0,
                    new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size, align='C')
    pdf.ln(line_height)
    
    ###### 날짜
    pdf.set_font('NanumGothic', '', 10)
    line_height = pdf.font_size * 2.5
    col_width = pdf.epw / 1
    pdf.multi_cell(col_width, line_height, f'{year}년 {month}월', border=0,
                    new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size, align='C')
    pdf.ln(line_height)

    ###### 소속 / 부서 / 사번 / 성명
    pdf.set_font('NanumGothic', '', 10)
    line_height = pdf.font_size * 2.5
    col_width = pdf.epw / len(info_cols)
    for c, v in zip(info_cols, uinfo):
        if c == 'No':
            c = '사번'
        pdf.multi_cell(col_width, line_height, f'{c} : {v}', border=0,
                        new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size, align='C')
    pdf.ln(line_height)

    ###### 지급내역 / 공제내역
    pdf.set_font('NanumGothicBold', '', 10)
    pdf.set_line_width(0.0)
    line_height = pdf.font_size * 2.5
    col_width = pdf.epw / 2
    # pdf.set_fill_color(50, 50, 50)
    pdf.multi_cell(col_width, line_height, '지급 내역', border=1,
                    new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                    align='C', fill=0)
    pdf.multi_cell(col_width, line_height, '공제 내역', border=1,
                    new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                    align='C')
    pdf.ln(line_height)


    ###### 정산내역(지급 / 공제) 디테일
    pdf.set_font('NanumGothic', '', 8)
    pdf.set_line_width(0.0)
    line_height = pdf.font_size * 2.5
    col_width = pdf.epw / 6
    for row in values:
        for val in row:
            pdf.multi_cell(col_width, line_height, val, border=1,
                            new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                            align='C')
        pdf.ln(line_height)


    ###### 지급 총액 / 공제 총액
    pdf.set_font('NanumGothicBold', '', 10)
    line_height = pdf.font_size * 2.5
    col_width = pdf.epw / 6
    sum_row = ['지급총액', format(pay_sum, ','), '', '공제총액', format(ded_sum, ','), '']
    for sr_i, sr in enumerate(sum_row):
        pdf.multi_cell(col_width, line_height, sr, border=1,
                        new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                        align='C')
    pdf.ln(line_height)

    ###### 실지급액
    pdf.set_font('NanumGothicBold', '', 10)
    pdf.set_line_width(0.0)
    line_height = pdf.font_size * 4
    col_width = pdf.epw / 4
    actual_pay_row = ['\n실지급액\n(지급총액 - 공제총액)', format(actual_pay, ','), '', '']
    for ap_i, ap_val in enumerate(actual_pay_row):
        is_border = ap_i < 2
        pdf.multi_cell(col_width, line_height, ap_val, border=is_border,
                        new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                        align='C')
    pdf.ln(line_height+4)

    ####### 유류대
    pdf.set_font('NanumGothicBold', '', 10)
    pdf.set_line_width(0.0)
    line_height = pdf.font_size * 2.5
    col_width = pdf.epw / 4
    oil_pay_row = ['유류대', format(oil_pay, ',')]
    for op_i, op_val in enumerate(oil_pay_row):
        pdf.multi_cell(col_width, line_height, op_val, border=1,
                        new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                        align='C')

    ####### 귀하의 노고에...
    col_width = pdf.epw / 2
    pdf.multi_cell(col_width, line_height, '* 귀하의 노고에 진심으로 감사드립니다 *', border=0,
                    new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                    align='C')
    pdf.ln(line_height)

    ####### 총지급액
    pdf.set_font('NanumGothicBold', '', 10)
    pdf.set_text_color(0,0,255)
    pdf.set_line_width(0.5)
    line_height = pdf.font_size * 4
    col_width = pdf.epw / 4
    total_pay_row = ['\n총지급액\n(실지급액 + 유류대)', format(total_pay, ',')]
    for tp_i, tp in enumerate(total_pay_row):
        pdf.multi_cell(col_width, line_height, tp, border=1,
                        new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                        align='C')
    ####### (주) 원로지스
    pdf.set_font('NanumGothicBold', '', 20)
    pdf.set_text_color(0,0,0)
    col_width = pdf.epw / 2
    pdf.multi_cell(col_width, line_height, '(주)   원   로   지   스', border=0,
                    new_x='RIGHT', new_y='TOP', max_line_height=pdf.font_size,
                    align='C')
    pdf.ln(line_height+3)
    

    # Save pdf
    usr_id, usr_name = uinfo[2:]
    pdf_save_path = f'{save_dir}/[{usr_id}]{usr_name}_{year}년{month}월_운송비정산내역.pdf'
    pdf.output(pdf_save_path)

    # # Save jpg
    # jpg_save_path = pdf_save_path.replace('.pdf', '.jpg')
    # pages = convert_from_path(pdf_save_path) # pdf to jpg
    # pages[0].save(jpg_save_path, "JPEG")

def remove_dir(dir_path):
    # 폴더, 파일 정리
    if os.path.exists(dir_path):
        shutil.rmtree(dir_path)

print("########## Run")

# ----- Streamlit
import streamlit as st
st.markdown("# 원로지스: 운송비 정산내역서 생성")
# file_path = "./(수정수정)고용산재추가_(주)원로지스 급여 22-11월.xls"

uploaded_file = st.file_uploader("**급여 엑셀파일 업로드**", type=['xls', 'xlsx'])
if uploaded_file is not None:
    if st.button('**정산내역서 만들기**'):
        with st.spinner("정산내역서 생성중..."):
            df, df_det, ym = preprocess_table(uploaded_file)
            year, month = ym.split('-')

            # 정산내역
            info_cols = ['소속', '부서', 'No', '성명']
            payment_cols = ['운송비', '세차비', '식대', '통행료', '휴무비용', '장거리수당', '고용산재보험', '추가배송비', '책임수당', '소급분명절수당', '지원내역(BRK+용품냉동)', '추가운행일수지원금', '휴대폰비용']
            deduction_cols = ['관리비', '협회비', '보험료', '고용산재지원분(사업자)', '고용산재지원분(기사)', '할부금', '환경개선부담금', '통신비', '공제1']

            replace_dict = {
                # "지급내역" 컬럼명 변경
                'P_Key': {
                    '추가배송비':'특별배송료',
                    '소급분명절수당':'소급, 수당분',
                    '지원내역(BRK+용품냉동)':'지원내역',
                    '휴대폰비용':'휴대폰요금지원금',
                    '0':'-',
                    0:'-'
                    },
                'P_Value': {'0':'-', 0:'-'},
                'P_Content':{'0':'-', 0:'-'},

                # "공제내역" 컬럼명 변경
                'D_Key' : {
                    '보험료':'차량보험료',
                    '공제1':'기타공제',
                    '0':'-',
                    0:'-',
                    },
                'D_Value': {'0':'-', 0:'-'},
                'D_Content':{'0':'-', 0:'-'}
            }

            save_dir = './tmp'
            zip_name = f'{year}년{month}월_운송비정산내역'

            # 폴더, 파일 정리
            remove_dir(save_dir)
            for z in glob('./*.zip'):
                os.remove(z)

            time.sleep(1)

            os.makedirs(save_dir, exist_ok=True)

            # 정산내역 파일 생성 시작
            users = sorted(list(df['No'].unique())) # 사번 정렬
            for uno in users:
                make_user_adjustment(uno, save_dir)
                # st.download_button('download pdf', data=bytes(user_adj.output()), file_name='test_file.pdf')

        st.info(f"총 {len(users)}명의 정산내역서가 생성되었습니다.", icon="🔔")

        # 압축, 다운로드 버튼
        shutil.make_archive(zip_name, 'zip', save_dir)
        with open(zip_name + '.zip', 'rb') as f:
            st.download_button('**다운로드**', data=f, file_name=zip_name + '.zip')

        remove_dir(save_dir) # pdf 파일 삭제
