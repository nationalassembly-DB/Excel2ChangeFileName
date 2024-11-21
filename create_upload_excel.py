"""main 함수"""


from module.change_filename import change_filename
from module.create_excel import create_excel


def main():
    """main 함수"""
    print('원하시는 작업을 선택하세요')
    select_input = input('1.엑셀 생성 2.파일명 변경 (번호 입력) 0.종료 : ')
    if select_input == '1':
        create_excel()
    elif select_input == '2':
        change_filename()
    elif select_input == '0':
        return
    else:
        print("다시 입력하세요\n")
        main()


if __name__ == "__main__":
    main()
