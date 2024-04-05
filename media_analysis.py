from analyze_media_keywords import analyze_media_keywords

# 분석할 파일들의 디렉토리와 파일명
directories = {
    '고발사주': ['고발사주_92.xlsx', '고발사주_393.xlsx', '고발사주.xlsx'],
    '도이치모터스_주가': ['도이치모터스_주가_김건희_Not-여사.xlsx', '도이치모터스_주가_조작_의혹.xlsx', '도이치모터스_주가_Not-김건희_여사.xlsx'],
    '양평': ['양평.xlsx', '양평_원희룡_김건희.xlsx', '양평_이재명_제외.xlsx', '양평_Not_원희룡_김건희.xlsx'],
    '이태원_특별법': ['이태원_특별법.xlsx', '이태원_특별법_거부권_행사.xlsx', '이태원_특별법_Not-거부권_행사.xlsx'],
    '해병대': ['해병대.xlsx', '해병대_윗선.xlsx', '해병대_제외_윗선.xlsx']
}

# 각 디렉토리와 파일에 대해 분석 수행
for directory, files in directories.items():
    for file in files:
        file_path = f'{directory}\\src\\{file}'
        output_file = analyze_media_keywords(file_path)
        print(f'분석 완료: {output_file}')
