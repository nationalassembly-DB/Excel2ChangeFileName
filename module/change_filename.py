"""
엑셀 파일을 사용하여 폴더명에 위치한 파일의 이름을 변경합니다
"""


import os
import pandas as pd


def change_filename():
    """
    엑셀파일을 사용하여 L열에 위치한 폴더명을 변경합니다
    엑셀파일 헤더 : YEAR,AUDITTYPE_CDB,COMMITTEE_ID,COMMITTEE_NAME,ORG_ID,
    ORG_NAME,PDF_NAME,HWP_NAME,PBM_NAME,BOOK_NAME,DIRECTORY_CDB,Path,FileName
    """
    print("\n>>>>>>파일명 변경<<<<<<\n")
    print("-"*24)
    excel_path = input("엑셀 파일 경로를 입력하세요: ")
    df = pd.read_excel(excel_path)

    old_names_dirname = df['Path'].tolist()  # L열 읽지 못함
    old_names_filename = df['FileName'].tolist()  # M열 읽지 못함
    old_names_with_dir = [os.path.join(dirname, filename) for dirname, filename in zip(
        old_names_dirname, old_names_filename)]

    new_names = df['PDF_NAME'].tolist()
    new_names = [os.path.splitext(name)[0] for name in new_names]
    extensions = ['PDF', 'HWP', 'PBM']

    for old_name, new_name, old_dirname in zip(old_names_with_dir, new_names, old_names_dirname):
        for ext in extensions:
            old_file = os.path.join("\\\\?\\", f"{old_name}.{ext}")
            new_file_with_dir = os.path.join("\\\\?\\",
                                             old_dirname, f"{new_name}.{ext}")
            if os.path.exists(old_file):
                os.rename(old_file, new_file_with_dir)
                print(f"파일명을 변경했습니다: {old_file} -> {new_file_with_dir}")
            else:
                print(f"파일이 존재하지 않습니다: {old_file}")

    print("작업이 완료되었습니다.")
