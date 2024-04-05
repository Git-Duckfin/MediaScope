import os
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

from scipy.cluster.hierarchy import linkage, leaves_list
import plotly.figure_factory as ff
import plotly.offline as py
import plotly.graph_objs as go
from plotly.subplots import make_subplots 

# Add the custom font
fe = fm.FontEntry(
    fname='NanumGothic.ttf',  # Update this path to the actual location of your font file
    name='NanumGothic')
fm.fontManager.ttflist.insert(0, fe)
plt.rcParams.update({'font.size': 18, 'font.family': 'NanumGothic'})
plt.rcParams['axes.unicode_minus'] = False

def create_heatmap_for_each_excel(directory):
    tab_figures = []
    file_titles = []

    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(directory, filename)
            df = pd.read_excel(file_path, usecols='A:K')
            
            heatmap_data = {}
            for index, row in df.iterrows():
                media = row.iloc[0]
                for col_index in range(1, len(row)):
                    keyword, value = eval(row.iloc[col_index])
                    if media not in heatmap_data:
                        heatmap_data[media] = {}
                    heatmap_data[media][keyword] = value

            heatmap_df = pd.DataFrame.from_dict(heatmap_data, orient='index')
            heatmap_df.replace([np.inf, -np.inf], np.nan, inplace=True)
            heatmap_df.fillna(0, inplace=True)

            row_linkage = linkage(heatmap_df.values, method='average')
            col_linkage = linkage(heatmap_df.T.values, method='average')
            row_order = leaves_list(row_linkage)
            col_order = leaves_list(col_linkage)
            clustered_df = heatmap_df.iloc[row_order, col_order]

            custom_colorscale = [
                [0.0, "red"],    # 낮은 값은 빨강색
                [0.5, "white"],  # 중간 값은 흰색
                [1.0, "blue"]    # 높은 값은 파랑색
            ]

            # 히트맵 생성 부분
            trace = go.Heatmap(
                z=clustered_df.values,
                x=clustered_df.columns[col_order],
                y=clustered_df.index[row_order],
                colorscale=custom_colorscale  # 커스텀 색상 스케일 적용
            )

            fig = go.Figure(data=[trace])
            tab_figures.append(fig)
            file_titles.append(filename)

    # 탭에 히트맵 추가
    fig = make_subplots(rows=1, cols=len(tab_figures), subplot_titles=file_titles)
    for i, tab_fig in enumerate(tab_figures, start=1):
        for trace in tab_fig['data']:
            fig.add_trace(trace, row=1, col=i)

    # HTML 파일로 저장
    output_dir = os.path.join(os.path.dirname(directory), 'graph')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    output_file = os.path.join(output_dir, 'heatmaps.html')
    py.plot(fig, filename=output_file, auto_open=False)

# 분석할 폴더 목록
words = ["고발사주", "도이치모터스_주가", "양평", "이태원_특별법", "해병대"]

for word in words:
    directory = os.path.join(word, 'out', 'analyzed')
    create_heatmap_for_each_excel(directory)

