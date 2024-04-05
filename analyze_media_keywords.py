import pandas as pd
import numpy as np
from collections import Counter
import os

def analyze_media_keywords(file_path):
    """
    언론사별 키워드 빈도수를 분석하고, 상위 10개 키워드 및 상관계수를 엑셀 파일로 저장하는 함수

    :param file_path: 분석할 엑셀 파일의 경로
    """
    # 엑셀 파일을 불러온 후 데이터 확인
    df = pd.read_excel(file_path)

    # '언론사'와 '키워드' 열을 추출
    media = df.iloc[:, 2]  # C열: '언론사'
    keywords = df.iloc[:, 14]  # O열: '키워드'

    # 언론사별 키워드 빈도수 계산
    media_keyword_freq = {}
    for m, k in zip(media, keywords):
        keyword_list = k.split(',')
        media_keyword_freq[m] = media_keyword_freq.get(m, Counter()) + Counter(keyword_list)

    # 각 언론사별 상위 40개 키워드 추출
    top_keywords_per_media = {m: freq.most_common(30) for m, freq in media_keyword_freq.items()}

    # 모든 언론사의 상위 키워드를 리스트로 변환
    all_top_keywords_list = list(set([kw[0] for keywords in top_keywords_per_media.values() for kw in keywords]))

    # 언론사별로 각 키워드의 빈도수를 나타내는 데이터 프레임 생성
    media_keyword_matrix = pd.DataFrame(index=media_keyword_freq.keys(), columns=all_top_keywords_list).fillna(0)
    for media, keywords in media_keyword_freq.items():
        for keyword, count in keywords.items():
            if keyword in all_top_keywords_list:
                media_keyword_matrix.at[media, keyword] = count

    # 상관계수 계산
    # correlation_matrix = media_keyword_matrix.corr()

    # 상위 디렉토리의 'out' 디렉토리 경로 생성
    out_dir = os.path.join(os.path.dirname(os.path.dirname(file_path)), 'out', 'freq')
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)

    # 'freq_'를 파일 이름 앞에 붙여서 저장할 파일 경로 지정
    base_file_name = os.path.basename(file_path)
    # 파일 이름에서 확장자 분리
    base_file_name, ext = os.path.splitext(base_file_name)

    # '_freq' 접미사를 확장자 전에 추가
    out_file_name = base_file_name + '_analyze' + ext
    out_file = os.path.join(out_dir, out_file_name)


    # 엑셀 파일 작성
    with pd.ExcelWriter(out_file) as writer:
        for media, keywords in top_keywords_per_media.items():
            df_top_keywords = pd.DataFrame(keywords, columns=['Keyword', 'Frequency'])
            df_top_keywords.to_excel(writer, sheet_name=media, index=False)
        # correlation_matrix.to_excel(writer, sheet_name='Correlation Matrix')

    return out_file

# 이 함수는 외부에서 파일 경로를 제공받아 해당 파일에 대한 분석을 수행하고, 결과를 'out' 폴더에 저장합니다.
