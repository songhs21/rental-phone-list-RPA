from tkinter import *
from openpyxl import load_workbook
from openpyxl.styles import Color, PatternFill
from datetime import datetime


########################################################################################

#Tkinter 생성
root = Tk() # Tkinter 생성
root.title("IMEI 추가 귀찮은ww") # 타이틀 설정
root.geometry("400x100") # 창 크기 설정
root.resizable(False, False) # 창 크기 변경 가능 여부 설정


# 엑셀 파일 불러오기
today = datetime.today()
rental = load_workbook("정합성 단말 대여 리스트_"+ str(today).replace("-","")[2:8] +".xlsx") # 엑셀파일에서 wb 호출
rentalSheet = rental.active # 활성화된 Sheet
db = load_workbook("list.xlsx")
dbSheet = db.active

########################################################################################

frame1 = Frame(root) # 프레임 생성
frame1.pack(side="left", fill="both", pady=(18,0), expand=True) # 프레임 표시

# IMEI
Label(frame1, text="IMEI").pack() # "이름" 라벨 생성 후 pack
num = Entry(frame1, width=30) # 텍스트 필드 생성 후 num 에 저장
num.pack() # num pack

########################################################################################

def Add():
    # 단말 대여 처리
    imei = num.get()
    Rimeis = []
    DBimeis = []
    if len(imei) == 15:

        for dx in dbSheet["E"]:
            DBimeis.append(dx.value) # list 엑셀에서 E열값 리스트 생성

        for rx in rentalSheet["F"]:
            Rimeis.append(rx.value) # 대여 리스트 엑셀에서 F열값 리스트 생성
        
        Eloc = (DBimeis.index(int(imei)))+1 # list 엑셀 E열값 중에서 입력받은 IMEI 값의 위치 값 추출

        CPinfo = [] # 입력받은 IMEI의 해당 단말 정보를 저장할 리스트
        for x in range(1, dbSheet.max_column+1):
            CPinfo.append(dbSheet.cell(row=Eloc, column=x).value) # list 엑셀에 입력받은 imei 단말 위치의 행 값을 추출
        
        
        if imei in Rimeis: # IMEI 값이 IMEIS에 들어있는지 체크
            print("이미 등록되어 있는 단말임")
        else:
            x= rentalSheet.max_row+1# 대여 리스트의 마지막 행 다음 행 위치 값을 고정하기 위한 변수
            n=0 # list 엑셀에서 추출한 데이터 리스트 순서를 체크할 변수
            for y in range(2, 7):
                rentalSheet.cell(row=x, column=y, value=str(CPinfo[n])) # list 엑셀에서 가져온 데이터를 대여 리스트 엑셀 마지막 행에 순차적으로 입력 
                n+=1
            rentalSheet.cell(row=rentalSheet.max_row, column=7, value="정합성/신뢰성")
            rentalSheet.cell(row=rentalSheet.max_row, column=9, value="대여")
            rentalSheet.cell(row=rentalSheet.max_row, column=10, value=str(today)[0:10])
            rentalSheet.cell(row=rentalSheet.max_row, column=11, value="O")
            rentalSheet.cell(row=rentalSheet.max_row, column=12, value="미반납")
            
            n=3 # 입력을 시작할 row 값
            for x in range(rentalSheet.max_row+1):
                rentalSheet.cell(row=n, column=1, value=x+1) # 3번째 줄 부터 No값 삽입
                n+=1 # 다음 행으로
                if n > rentalSheet.max_row:
                    break # n 값이 마지막 행을 넘어갈 경우 반복문 종료
            
            rental.save("정합성 단말 대여 리스트_"+ str(today).replace("-","")[2:8] +".xlsx") # IMEI 추가 후 저장
        
        # 정합성 대여 단말 리스트에 단말 추가 끝

    else:
        print("imei 자리 수 확인 필요")
########################################################################################


########################################################################################
# 단말 반납 처리
def back():
    imei = num.get()
    Rimeis = []

    for rx in rentalSheet["F"]:
        Rimeis.append(rx.value) # 대여 리스트 엑셀에서 F열값 리스트 생성

    Floc = (Rimeis.index(int(imei)))+1

    if int(imei) in Rimeis : # IMEI 값이 RIMEIS에 들어있는지 체크
        for x in range(1, rentalSheet.max_column+1):
            # 반납에 해당하는 컬러 값"d9d9d9"
            rentalSheet.cell(row=Floc, column=x).fill = PatternFill(start_color="d9d9d9", end_color="d9d9d9", fill_type="solid")
            rentalSheet.cell(row=Floc, column=12, value="반납")
            rentalSheet.cell(row=Floc, column=13, value="")
        rental.save("정합성 단말 대여 리스트_"+ str(today).replace("-","")[2:8] +".xlsx") # IMEI 추가 후 저장
    else:
        print("없는데 뭘 찾는 거야")
########################################################################################


frame2 = Frame(root)
frame2.pack(side="right", fill="both", expand=True)
Button(frame2, text="대여", command=Add).pack(fill="both", expand=True)
Button(frame2, text="반납", command=back).pack(fill="both", expand=True)

root.mainloop()
