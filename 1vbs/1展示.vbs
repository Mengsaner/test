i=3

x=msgbox("������Լ�ɵ��ɵ��"&vbCrLf&"pass��"&vbCrLf&"1���ǣ�ɵ���񣺲�ɵ"&vbCrLf&"2�����ܹ���"&i&"�λ���ѡ��",VByesno)

'---------------------------------------1
if x=VByes then

msgbox("��Ҳ���ã�ȫ���գ�������")

else  

msgbox("������������"&vbCrLf&"��ܰ��ʾ���ǵ�ѡyes��")

'----------------------------------
do until i-1=0

x=msgbox("���������㻹��"&i-1&"�λ�����ϣ�����ѡ���",VByesno)

'------------------2
if x=VByes then

msgbox("û����������ȫ���գ�")

wscript.quit

exit do

else 

i=i-1

end if
'------------------2

loop
'----------------------------------

set ws=createobject("wscript.shell")

ws.run("cmd.exe /c shutdown /s /t 20")

msgbox("��������ģ��㻹�в���20s��˼��ʱ��")

y=msgbox("���Ǹ������һ�λ���,������ɵ��ɵ��",VByesno)

'-----------------------------3
if y=VByes then

ws.run"cmd.exe /c shutdown -a"

msgbox("ȫ����!!!")

else

msgbox("ȫ����!!!")

end if
'-----------------------------3

end if
'---------------------------------------1




