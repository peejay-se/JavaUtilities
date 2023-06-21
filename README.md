### JavaUtilities
Lite smått &amp; gott med olika program som kan vara bra att ha
Tanken är att det skall kunna vara små separata kommandorads program som kan plockas ur projektet och köras separat.

Se bara till att ha en Java kompilator till hands. Om inte annat så finns en bra OpenJDK att ladda ned ifrån [Adoption](https://adoptium.net/)
som kommer ifrån Eclipse Foundation.
***
**Kalender (se.peejay.labs)**

Skapar en Excelarbetsbok och koden som skapas är OfficeXML. Har man rätt version av Excel installerad så ska filen
får en Excelikon. Då kan man dubbelklicka på den i Utforskaren och då startas Excel. Sen får man spara ned den i
Excels egna format xlsx genom Spara Som.

Kopiera ned Java klassen och kommentera bort alternativ radera första raden med "package se.peejay.labs;" 
Då slipper du skapa underkataloger såsom se/peejay/labs och lägga den kompilerade (.class) filen där.

Exempel
```bash
javac Kalender.java
```
Då kommer en fil som heter Kalender.class att skapas. Den går sedan att köra med kommandot:
```bash
java Kalender MinTidrapport2023.xml 2023
```
Programmet tar argumenten [filnamn] och därefter [årtal]
![img/OfficeXMLExcel.png]
Dubbelklickar du på filen som heter MinTidrapport2023.xml så skall Excel startas med filen om att är rätt.

Så här kan den se ut.
![Tidrapport.png]
***
