from analyze_and_export_keywords import analyze_and_export_keywords

# media_analysis.py 코드를 실행
file_path = 'media_analysis.py'

# 파일 읽기
# UTF-8 인코딩을 사용하여 파일 읽기
with open(file_path, 'r', encoding='utf-8') as file:
    media_analysis_code = file.read()

# 코드 실행
exec(media_analysis_code)

# 분석할 디렉토리 이름 설정
words = ["고발사주", "도이치모터스_주가", "양평", "이태원_특별법", "해병대"]

# 함수 호출
# analyze_and_export_keywords(word)
for word in words:
    analyze_and_export_keywords(word)

# create_heatmap_from_excel.py 코드를 실행
file_path = 'create_heatmap_from_excel.py'

# 파일 읽기
# UTF-8 인코딩을 사용하여 파일 읽기
with open(file_path, 'r', encoding='utf-8') as file:
    create_heatmap_code = file.read()

# 코드 실행
exec(create_heatmap_code)
