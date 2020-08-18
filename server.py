import socket
from win32com.client import Dispatch

def speak(str1):
	speak=Dispatch(("SAPI.SpVoice"))
	speak.speak(str1)

s=socket.socket(socket.AF_INET,socket.SOCK_STREAM,0)
port=4702
s.bind((socket.gethostname(),port))
s.listen(5)
conn,addr=s.accept()
print(f"Got Connected to {addr}")
speak(f"Got Connected to {addr}")

while True:
	data=input("Enter Your Message: ")
	conn.send(bytes(data,encoding='utf-8'))
	print(conn.recv(1024))
	speak("New Message From Client")
