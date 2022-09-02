C:\Program Files\LibreOffice\program>soffice "--accept=socket,host=127.0.0.1,port=2002,tcpNoDelay=1;urp;" --writer --norestore --headless

izgleda da da port treba ipak biti 8100


Py UNO je problematičan jer traži da se pokrene py koji je isporučen unutar LO-a.
Npr LO 7.4 sadrži py3.8 i .pyd

c:\Program Files\LibreOffice\program\pyuno.pyd

ovaj dll je linkan sa python38.dll tako da se ne moze koristiti u npr. py 3.10

Zato je najbolje koristit java uno

== Java UNO ===

= compile =
c:\dev
javac -cp "c:/Program Files/LibreOffice/program/classes/libreoffice.jar SimpleBootstrap_java.java

= run =
mora se pokrentuti iz LO/programs

cd c:\Program Files\LibreOffice\program
java -cp classes/libreoffice.jar;c:/dev SimpleBootstrap_java


C:\dev\ziher\ziher_mono\lo_uno_py_java\java>
set LO_CLASSES="c:/dev/ziher/ziher_mono/lo_uno_py_java/java;c:/Program Files/LibreOffice/program/classes/libreoffice.jar"

// utf-8 encode string (naslov grafikona)
javac -encoding "UTF-8" -cp %LO_CLASSES% SCalc.java

set LO_PROGRAM="c:\Program Files\LibreOffice\program"
cd %LO_PROGRAM%

REM set LO_CLASSES="c:/Program Files/LibreOffice/program/classes/libreoffice.jar"

java -Dfile.encoding=UTF-8 -cp %LO_CLASSES% SCalc
