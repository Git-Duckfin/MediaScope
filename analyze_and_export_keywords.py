import pandas as pd
import os

def analyze_and_export_keywords(word):
    # 디렉토리 경로 설정
    base_dir = word
    freq_dir = os.path.join(base_dir, 'out', 'freq')

    # 디렉토리가 존재하지 않으면 생성
    if not os.path.exists(freq_dir):
        os.makedirs(freq_dir)

    # freq_dir 디렉토리 내의 모든 엑셀 파일 처리
    for file_name in os.listdir(freq_dir):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(freq_dir, file_name)
            output_file = process_excel_file(file_path)
            print(f'Processed: {file_path}\nOutput: {output_file}')

def process_excel_file(file_path):
    # 파일을 읽어 각 시트별로 키워드 빈도 계산
    with pd.ExcelFile(file_path) as xls:
        media_keyword_freq = {}
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            media_keyword_freq[sheet_name] = df.set_index('Keyword')['Frequency'].to_dict()

    # 전체 키워드 빈도 계산 및 상위 20개 추출
    total_keyword_freq = pd.Series({k: sum(d.get(k, 0) for d in media_keyword_freq.values()) for k in set().union(*media_keyword_freq.values())})
    top_20_keywords = total_keyword_freq.nlargest(20).index

    # 각 언론사별로 키워드의 빈도 비중 계산
    top_keywords_matrix = {}
    for media, keywords in media_keyword_freq.items():
        total = sum(keywords.values())
        proportions = {k: keywords.get(k, 0) / total for k in top_20_keywords}
        top_keywords_matrix[media] = sorted(proportions.items(), key=lambda x: x[1], reverse=True)[:10]

    # 결과 데이터프레임 생성
    result_df = pd.DataFrame.from_dict(top_keywords_matrix, orient='index', columns=[f'Top {i+1}' for i in range(10)])

    # 파일 저장 경로 설정 및 결과 저장
    output_file_path = export_to_excel(result_df, file_path)
    return output_file_path

def export_to_excel(dataframe, original_file_path):
    # 파일 저장 경로 설정
    output_dir = os.path.join(os.path.dirname(os.path.dirname(original_file_path)), 'analyzed')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    output_file_name = os.path.basename(original_file_path)
    output_file_path = os.path.join(output_dir, output_file_name)

    # 결과를 엑셀 파일로 저장
    with pd.ExcelWriter(output_file_path) as writer:
        dataframe.to_excel(writer, sheet_name='Top Keywords Proportion')

    return output_file_path
