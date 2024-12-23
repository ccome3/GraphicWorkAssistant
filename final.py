import os
import pyautogui
import ctypes
import subprocess
from datetime import datetime
from tkinter import *
from tkinter import filedialog
from time import sleep
import threading

# 스크린샷 저장 경로와 기타 설정 변수
base_dir = ""
screenshot_interval = 10 * 60  # 기본값: 10분
light_mode_hour = 10  # 기본값: 10시
light_mode_minute = 0  # 기본값: 00분
dark_mode_hour = 18  # 기본값: 18시
dark_mode_minute = 0  # 기본값: 00분

# 다크 모드 / 라이트 모드 설정 함수 (PowerShell 사용)
def set_dark_mode(enable=True):
    """PowerShell을 통해 다크 모드 또는 라이트 모드를 설정하는 함수"""
    if enable:
        # 다크 모드로 설정
        command = 'Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize" -Name "AppsUseLightTheme" -Value 0'
    else:
        # 라이트 모드로 설정
        command = 'Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize" -Name "AppsUseLightTheme" -Value 1'

    try:
        # PowerShell 명령어 실행
        subprocess.run(["powershell", "-Command", command], check=True)
        print(f"모드 {'다크' if enable else '라이트'}로 전환 완료")
    except subprocess.CalledProcessError as e:
        print(f"모드 전환 실패: {e}")

# 스크린샷 캡처 함수
def capture_screenshot():
    """스크린샷을 캡처하고 저장하는 함수"""
    if not base_dir:
        print("경로가 설정되지 않았습니다.")
        return

    current_time = datetime.now()
    date_folder = current_time.strftime("%Y-%m-%d")  # '2024-12-21' 형태
    timestamp = current_time.strftime("%H-%M-%S")  # '10-30-15' 형태

    # 날짜별 폴더 경로 생성
    folder_path = os.path.join(base_dir, date_folder)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    # 스크린샷 파일 경로
    screenshot_filename = f"{timestamp}.png"
    screenshot_path = os.path.join(folder_path, screenshot_filename)

    # 화면 캡처 및 저장
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_path)
    print(f"스크린샷 저장됨: {screenshot_path}")

# 작업 흐름 관리 함수
def work_flow_manager():
    """사용자가 설정한 대로 작업 흐름을 관리"""
    while True:
        # 현재 시각을 가져옴
        current_time = datetime.now()

        # 다크 모드 전환 시각 체크
        if current_time.hour == dark_mode_hour and current_time.minute == dark_mode_minute:
            set_dark_mode(True)

        # 라이트 모드 전환 시각 체크
        if current_time.hour == light_mode_hour and current_time.minute == light_mode_minute:
            set_dark_mode(False)

        # 주기적으로 스크린샷 캡처
        capture_screenshot()

        sleep(screenshot_interval)

# 폴더 선택 함수
def select_folder():
    """사용자가 폴더를 선택하는 함수"""
    global base_dir
    base_dir = filedialog.askdirectory()  # 폴더 선택 다이얼로그
    folder_label.config(text=f"선택된 경로: {base_dir}")

# 시작 버튼을 눌렀을 때 호출되는 함수
def start_work():
    """시작 버튼 클릭 시 작업 흐름을 시작하는 함수"""
    global screenshot_interval, light_mode_hour, light_mode_minute, dark_mode_hour, dark_mode_minute

    # 사용자 입력값을 가져와 설정
    try:
        screenshot_interval = int(screenshot_interval_entry.get()) * 60  # 분 단위로 변환
        light_mode_hour = int(light_mode_hour_combobox.get())
        light_mode_minute = int(light_mode_minute_combobox.get())
        dark_mode_hour = int(dark_mode_hour_combobox.get())
        dark_mode_minute = int(dark_mode_minute_combobox.get())

        print(f"스크린샷 간격: {screenshot_interval // 60}분, 라이트 모드: {light_mode_hour}:{light_mode_minute}, 다크 모드: {dark_mode_hour}:{dark_mode_minute}")
        
        # 현재 시각 확인하여 모드 설정
        current_time = datetime.now()

        # 현재 시간이 라이트모드 시간 범위에 포함되면 라이트모드로 설정
        if (current_time.hour > dark_mode_hour or (current_time.hour == dark_mode_hour and current_time.minute >= dark_mode_minute)) and \
           (current_time.hour < light_mode_hour or (current_time.hour == light_mode_hour and current_time.minute < light_mode_minute)):
            set_dark_mode(False)  # 라이트 모드로 설정
        else:
            set_dark_mode(True)  # 다크 모드로 설정

        # 새로운 인터페이스를 띄움
        start_new_interface()

        # 기존 인터페이스 종료는 새로운 인터페이스가 다 띄운 후에 진행
        root.after(500, root.quit)  # 500ms 후에 기존 창 종료

    except ValueError:
        print("잘못된 입력값이 있습니다. 숫자만 입력하세요.")

# 새로운 작업 인터페이스 생성
def start_new_interface():
    """작업 흐름을 관리하는 새로운 인터페이스를 생성하는 함수"""
    new_window = Toplevel(root)
    new_window.title("작업 흐름 관리 (작업 중)")

    # 즉시 스크린샷 찍기 버튼
    capture_button = Button(new_window, text="즉시 스크린샷 찍기", command=capture_screenshot)
    capture_button.pack(pady=20)

    # 프로그램 종료 버튼
    exit_button = Button(new_window, text="프로그램 종료", command=root.quit)
    exit_button.pack(pady=20)

    # 새로운 인터페이스에서 작업 흐름 관리
    threading.Thread(target=work_flow_manager, daemon=True).start()

    # 대기상태에서 창을 닫지 않도록 유지
    new_window.mainloop()

# Tkinter GUI 설정
root = Tk()
root.title("Graphic Work Assistant")

# 폴더 선택 레이블과 버튼
folder_label = Label(root, text="경로를 선택하세요.", width=50, anchor='w')
folder_label.grid(row=0, column=0, padx=10, pady=10)

folder_button = Button(root, text="폴더 선택", command=select_folder)
folder_button.grid(row=0, column=1, padx=10, pady=10)

# 스크린샷 간격 입력
interval_label = Label(root, text="스크린샷 간격 (분):")
interval_label.grid(row=1, column=0, padx=10, pady=10)

screenshot_interval_entry = Entry(root)
screenshot_interval_entry.grid(row=1, column=1, padx=10, pady=10)
screenshot_interval_entry.insert(0, "10")  # 기본값 10분

# 라이트 모드 시각 입력
light_mode_label = Label(root, text="라이트 모드 전환 시각 (시:분):")
light_mode_label.grid(row=2, column=0, padx=10, pady=10)

# 라이트 모드 시/분 선택
light_mode_hour_combobox = Spinbox(root, from_=0, to=23, width=3)
light_mode_hour_combobox.grid(row=2, column=1, padx=10, pady=10)
light_mode_hour_combobox.delete(0, END)  # 기존 값 삭제
light_mode_hour_combobox.insert(0, "10")  # 기본값 10시

light_mode_minute_combobox = Spinbox(root, from_=0, to=59, width=3)
light_mode_minute_combobox.grid(row=2, column=2, padx=10, pady=10)
light_mode_minute_combobox.delete(0, END)  # 기존 값 삭제
light_mode_minute_combobox.insert(0, "00")  # 기본값 00분

# 다크 모드 시각 입력
dark_mode_label = Label(root, text="다크 모드 전환 시각 (시:분):")
dark_mode_label.grid(row=3, column=0, padx=10, pady=10)

# 다크 모드 시/분 선택
dark_mode_hour_combobox = Spinbox(root, from_=0, to=23, width=3)
dark_mode_hour_combobox.grid(row=3, column=1, padx=10, pady=10)
dark_mode_hour_combobox.delete(0, END)  # 기존 값 삭제
dark_mode_hour_combobox.insert(0, "18")  # 기본값 18시

dark_mode_minute_combobox = Spinbox(root, from_=0, to=59, width=3)
dark_mode_minute_combobox.grid(row=3, column=2, padx=10, pady=10)
dark_mode_minute_combobox.delete(0, END)  # 기존 값 삭제
dark_mode_minute_combobox.insert(0, "00")  # 기본값 00분

# 시작 버튼
start_button = Button(root, text="작업 시작", command=start_work)
start_button.grid(row=4, column=0, columnspan=3, pady=20)

# Tkinter GUI 루프 시작
root.mainloop()
