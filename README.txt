C:\Program Files\LibreOffice\program>soffice "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;" --writer --norestore --headless

izgleda da da port treba ipak biti 8100


Py UNO je problematičan jer traži da se pokrene py koji je isporučen unutar LO-a.
Npr LO 7.4 sadrži py3.8 i .pyd

c:\Program Files\LibreOffice\program\pyuno.pyd

ovaj dll je linkan sa python38.dll tako da se ne moze koristiti u npr. py 3.10

Zato je najbolje koristit java uno