#imports
import subprocess
import customtkinter as ctk
from customtkinter import CTkLabel, CTkSwitch, CTkButton
from tkinter import messagebox
import os
from dotenv import load_dotenv, set_key
import sys
import ctypes


#runtime
load_dotenv()
cortana_dis = os.getenv("CORTANA_DIS", "False").lower() == "true"
copilot_dis = os.getenv("COPILOT_DIS", "False").lower() == "true"
gamebar_dis = os.getenv("GAMEBAR_DIS", "False").lower() == "true"
news_dis = os.getenv("NEWS_DIS", "False").lower() == "true"
command = ""

#check admin
def is_admin():
    return ctypes.windll.shell32.IsUserAnAdmin()
is_admin()
if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

#root creation
root = ctk.CTk()

#root configuration
root.iconbitmap("icon.ico")
root.configure(fg_color = '#5A9BD4')
root.geometry('300x400')
root.title("WinTrim")

#functions
def dis_cortana():
    global cortana_dis
    if cortana_dis:
        print("Cortana enabled")
        cortana_dis = False
    else:
        print("Cortana disabled")
        cortana_dis = True

def dis_copilot():
    global copilot_dis
    if copilot_dis:
        print("Copilot enabled")
        copilot_dis = False
    else:
        print("Copilot disabled")
        copilot_dis = True

def dis_gamebar():
    global gamebar_dis
    if gamebar_dis:
        print("Game bar enabled")
        gamebar_dis = False
    else:
        print("Game bar disabled")
        gamebar_dis = True

def dis_news():
    global news_dis
    if news_dis:
        print("News enabled")
        news_dis = False
    else:
        print("News disabled")
        news_dis = True

def save():
    global command
    print("")
    print("Settings saved:")
    if cortana_dis:
        command = subprocess.run("""Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name AllowCortana -Value 0""", shell = True, capture_output = True, text = True)
        print(f"Cortana disabled - command output: {command}")
        set_key(".env", "CORTANA_DIS", "True")
    elif not cortana_dis:
        command = subprocess.run("""Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name AllowCortana -Value 1""", shell = True, capture_output = True, text = True)
        print(f"Cortana enabled - command output: {command}")
        set_key(".env", "CORTANA_DIS", "False")

    if copilot_dis:
        command = subprocess.run("""reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /t REG_DWORD /d 1 /f""", shell = True, capture_output = True, text = True)
        print(f"Copilot disabled - command output: {command}")
        set_key(".env", "COPILOT_DIS", "True")
    elif not copilot_dis:
        command = subprocess.run("""reg delete "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot" /v TurnOffWindowsCopilot /f""", shell = True, capture_output = True, text = True)
        print(f"Copilot enabled - command output: {command}")
        set_key(".env", "COPILOT_DIS", "False")

    if gamebar_dis:
        command = subprocess.run("""reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f""", shell = True, capture_output = True, text = True)
        command = subprocess.run("""reg add "HKCU\Software\Microsoft\GameBar" /v AllowAutoGameMode /t REG_DWORD /d 0 /f""",shell=True, capture_output=True, text=True)
        command = subprocess.run("""reg add "HKLM\SYSTEM\CurrentControlSet\Services\XblGameSave" /v Start /t REG_DWORD /d 4 /f""",shell=True, capture_output=True, text=True)
        command = subprocess.run("""reg add "HKLM\SYSTEM\CurrentControlSet\Services\XblAuthManager" /v Start /t REG_DWORD /d 4 /f""", shell=True,capture_output=True, text=True)
        print(f"Game bar disabled - command output: {command}")
        set_key(".env", "GAMEBAR_DIS", "True")
    elif not gamebar_dis:
        command = subprocess.run("""reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 1 /f""",shell=True, capture_output=True, text=True)
        command = subprocess.run("""reg add "HKCU\Software\Microsoft\GameBar" /v AllowAutoGameMode /t REG_DWORD /d 1 /f""", shell=True,capture_output=True, text=True)
        command = subprocess.run("""reg add "HKLM\SYSTEM\CurrentControlSet\Services\XblGameSave" /v Start /t REG_DWORD /d 3 /f""",shell=True, capture_output=True, text=True)
        command = subprocess.run("""reg add "HKLM\SYSTEM\CurrentControlSet\Services\XblAuthManager" /v Start /t REG_DWORD /d 3 /f""", shell=True,capture_output=True, text=True)
        print(f"Game bar enabled - command output: {command}")
        set_key(".env", "GAMEBAR_DIS", "False")

    if news_dis:
        command = subprocess.run("""reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v TaskbarDa /t REG_DWORD /d 0 /f""",shell=True, capture_output=True, text=True)
        print(f"News disabled - command output: {command}")
        set_key(".env", "NEWS_DIS", "True")
    elif not news_dis:
        command = subprocess.run("""reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" /v TaskbarDa /t REG_DWORD /d 1 /f""", shell=True,capture_output=True, text=True)
        print(f"News enabled - command output: {command}")
        set_key(".env", "NEWS_DIS", "False")

    msgbox = messagebox.askokcancel("Restart PC?", "To apply the changes, please restart your PC.")

    if msgbox:
        print("PC restarting...")
        subprocess.run("shutdown /r /f /t 0", shell=True)
    if not msgbox:
        print("changes will be applied on restart")
        messagebox.showinfo("Changes not applied",  "Changes will be applied on restart")

#widgets
home_heading = CTkLabel(
    root,
    text_color='black',
    fg_color='#5A9BD4',
    font=('Exo 2 Extra Bold', 70),
    text="WinTrim"
)


home_desc = CTkLabel(
    root,
    text_color = 'black',
    fg_color = '#5A9BD4',
    font = ('Exo 2 Extra Bold', 17),
    text = "All toggled options will be disabled:"
)
cortana = CTkSwitch(
    root,
    text = "Cortana",
    command = dis_cortana,
    font =('Exo 2 Extra Bold', 20),
    text_color = 'black'
)
if cortana_dis:
    cortana.select()

copilot = CTkSwitch(
    root,
    text = "Copilot (Windows 11)",
    command = dis_copilot,
    font =('Exo 2 Extra Bold', 20),
    text_color = 'black'
)
if copilot_dis:
    copilot.select()

gamebar = CTkSwitch(
    root,
    text = "Game Bar",
    command = dis_gamebar,
    font =('Exo 2 Extra Bold', 20),
    text_color = 'black'
)
if gamebar_dis:
    gamebar.select()

news = CTkSwitch(
    root,
    text = "News",
    command = dis_news,
    font =('Exo 2 Extra Bold', 20),
    text_color = 'black'
)
if news_dis:
    news.select()

save = CTkButton(
    root,
    text = "Save",
    font = ('Exo 2 Extra Bold', 17),
    text_color = 'black',
    command = save
)

#packing order
home_heading.pack()
home_desc.pack()
news.pack(pady = 10)
cortana.pack(pady = 10)
gamebar.pack(pady = 10)
copilot.pack(pady = 10)
save.pack(pady = 10)

#root mainloop
root.mainloop()
