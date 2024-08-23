import pandas as pd
import os

# 동일 파일명 handling 필요
#! L열에 파일명 대신 폴더 경로 사용하는 것으로 확정


def change_filename(excel_file_path):
    df = pd.read_excel(excel_file_path)

    old_names = df['L'].tolist()
    new_names_ext = df['G'].tolist()
    new_names = [os.path.splitext(name)[0] for name in new_names_ext]
    extensions = ['PDF', 'HWP', 'PBM']

    for old_name, new_name in zip(old_names, new_names):
        for ext in extensions:
            old_file = os.path.join(f"{old_name}.{ext}")
            new_file = os.path.join(f"{new_name}.{ext}")
            if os.path.exists(old_file):
                os.rename(old_file, new_file)
                print(f"파일명을 변경했습니다: {old_file} -> {new_file}")
            else:
                print(f"파일이 존재하지 않습니다: {old_file}")


if __name__ == "__main__":
    excel_file_path = input("엑셀 파일 경로를 입력하세요: ")
    change_filename(excel_file_path)
    print("작업이 완료되었습니다.")
