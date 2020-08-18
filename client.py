import socket
from win32com.client import Dispatch

def speak(str1):
	speak=Dispatch(("SAPI.SpVoice"))
	speak.speak(str1)

s=socket.socket(socket.AF_INET,socket.SOCK_STREAM,0)
port=4702
s.connect((socket.gethostname(),port))

try:
	while True:
		print(s.recv(1024))
		speak("New Message From Server")
		data=input("Enter Your Message: ")
		s.send(bytes(data,encoding='utf-8'))
except Exception as e:
	speak("Connection Lost")

