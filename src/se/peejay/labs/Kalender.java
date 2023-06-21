package se.peejay.labs;

//======================================================================================================================================
//                                   Skapa en årskalender för tidrapportering i Excel
//======================================================================================================================================
//  Skapad:   2023-06-20
//  Filnamn:  Kalender.java
//======================================================================================================================================
//   Skapar en fil med OfficeXML för att kunna användas som exempelvis en tidrapporteringsmall.
//   Filen som skapas kommer att se ut som en Excelfil (om installationen är rätt registrerad) i utforskaren om man tittar på den.
//   Man kan få en varning när filen öppnas, men det beror på att den innehåller formulas som beräknas. 
//
//   Man måste ha Java installerat på maskinen. Ladda ned filen och öppna en kommandoprompt och skriv:
//   javac Kalender.java
//
//   Filen kommer att kompileras och det skapas en fil som heter Kalender.class. Skapa underkatalogerna se/peejay/labs och lägg
//   class filen i labs-katalogen. Vill man slippa detta så får man kommentera bort/radera första raden package se.peejay.labs i koden.
//
//   Sen kan man skriva:
//   java Kalender.class MinKalender.xml 2024
//
//   Om allt gått bra så ska du ha en fil som heter MinKalender.xml som du ska kunna dubbelklicka på i Utforskaren och Excel skall
//   startas med filen. Därefter får man spara den som exempelvis MinKalender.xlsx
//
//   Argumenten för Kalender är Kalender.class [filnamn] [årtal]
//
//   För man något fel i koden så är det inte så lätt att debugga. Felmeddelandet ifrån Excel säger inte speciellt mycket
// 
//======================================================================================================================================
//  Revisionshistorik
//======================================================================================================================================
//
//======================================================================================================================================

import java.io.FileWriter;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

// ======================================================================================================================================
//
// ======================================================================================================================================
public class Kalender {

  static FileWriter officeXMLFile = null;
  static Calendar cal = null;

  // ======================================================================================================================================
  //   Huvudprogrammet
  // ======================================================================================================================================
  public static void main(String[] args) {

    // ====[ Kontrollera argumenten ]====
    if (args.length != 2) {
      System.out.println("Fel antal argument....");
      System.out.println("Kalender [filnamn] [årtal]");
      System.exit(1);
    }

    int veckoStart;

    try {

      cal = Calendar.getInstance();

      if (!isNumeric(args[1])) {
        System.out.println("Felaktigt årtal!");
        System.exit(0);
      }

      // ====[ Bygg upp den första dagen på det angivna årtalet exempelvis 2020-01-01
      StringBuffer sb = new StringBuffer(args[1]);
      sb.append("-01-01");
      
      SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
      Date dateObj = sdf.parse(sb.toString());
      cal.setTime(dateObj);

      
      // The resulting number ranges from 1 (Sunday) to 7 (Saturday).
      // The first month of the year in the Gregorian and Julian calendars is JANUARY
      // which is 0

      // 1 = Söndag, 2 = Måndag, 3 = Tisdag, 4 = Onsdag, 5 = Torsdag, 6 = Fredag, 7 = Lördag
      
      // ====[ Räkna ut vilken dag som är måndagen, har alltid hela veckor ]====
      if (cal.get(Calendar.DAY_OF_WEEK) == 1) {
        veckoStart = -6;
      } else {
        veckoStart = (cal.get(Calendar.DAY_OF_WEEK) - 2) * -1;
      }

      cal.add(Calendar.DATE, veckoStart);

      officeXMLFile = new FileWriter(args[0]);

      // ====[ Skriv ut början på filen ]====
      WriteFileStart();

      // ====[ Dokumentegenskaper, Dokumentinställningar, Formattering ]====
      WriteDocumentProperties();
      WriteOfficeDocumentSettings();
      WriteExcelWorkbook();
      WriteStyles();

      // ====[ Skriv ut ]====
      WriteWorksheet();

      WriteWorkSheetOptions();
      WriteFileEnd();
      officeXMLFile.close();
      
      System.out.println("\n\nKalendern är klar..");
    } catch (IOException | ParseException e) {
      System.out.println("Ett fel uppstod");
      e.printStackTrace();
    }

  }

  // ======================================================================================================================================
  //   Skriv kalkylbladet som heter Kalender
  // ======================================================================================================================================
  private static void WriteWorksheet() throws IOException {
    
    int rowNo=0;
    
    officeXMLFile.write("<Worksheet ss:Name=\"Kalender\">");  // Namnet på fliken i arbetsboken
    
    // ====[ ExpandedRowCount styr hur många rader som skall visas ]====
    officeXMLFile.write("<Table ss:ExpandedColumnCount=\"12\" ss:ExpandedRowCount=\"600\" x:FullColumns=\"1\" x:FullRows=\"1\" ss:StyleID=\"s62\">");
    officeXMLFile.write("<Column ss:StyleID=\"s62\" ss:AutoFitWidth=\"0\" ss:Width=\"23.25\"/>");
    officeXMLFile.write("<Column ss:StyleID=\"s62\" ss:Width=\"48.75\"/>");
    officeXMLFile.write("<Column ss:StyleID=\"s62\" ss:Width=\"55.5\"/>");
    officeXMLFile.write("<Column ss:StyleID=\"s62\" ss:Width=\"53.25\"/>");
    officeXMLFile.write("<Column ss:Index=\"9\" ss:StyleID=\"s62\" ss:Width=\"48.75\" ss:Span=\"1\"/>");
    officeXMLFile.write("<Column ss:Index=\"12\" ss:StyleID=\"s62\" ss:AutoFitWidth=\"0\" ss:Width=\"405.75\"/>");
    
    // ====[ Loopa igenom veckorna ]====
    for(int vecka=1;vecka<55;vecka++) {

      // ====[ Skriv ut huvudet för veckan ]====
      WriteTableHeader(vecka);
      rowNo++;
      
      System.out.printf("Skriver ut ett veckoavsnitt (%s avsnitt)\n",vecka);
      
      // =====[ Loopa igenom alla veckodagar ]====
      for(int dag=1;dag<8;dag++) {

        // ====[ Skriv ut en rad för varje dag ]====
        WriteDayData(vecka,dag);  

        rowNo++;
        cal.add(Calendar.DATE, 1);
      }
      
      WriteEndOfDayData();
      rowNo++;
      rowNo++;
    }


    StringBuffer sb = new StringBuffer();
    officeXMLFile.write(sb.toString());
    
    
    officeXMLFile.write("</Table>");

  }

  // ======================================================================================================================================
  //   Skriv ut veckosummeringen och lägg en rad innan nästa avsnitt
  // ======================================================================================================================================
  private static void WriteEndOfDayData() throws IOException {
    
    officeXMLFile.write("<Row>");
    officeXMLFile.write("<Cell ss:Index=\"2\" ss:StyleID=\"s76\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s77\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s77\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s76\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s76\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s76\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s78\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s79\" ss:Formula=\"=SUM(R[-7]C:R[-1]C)\">"); // Summerar veckans verkliga tid ( beräknat på minuten )
    officeXMLFile.write("<Data ss:Type=\"DateTime\" />");
    officeXMLFile.write("</Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"s80\" ss:Formula=\"=SUM(R[-7]C:R[-1]C)\">"); // Summerar den tid som du skrivit i tidskolumnen
    officeXMLFile.write("<Data ss:Type=\"Number\">0</Data>");
    officeXMLFile.write("</Cell>");
    officeXMLFile.write("<Cell ss:Index=\"12\" ss:StyleID=\"s81\"/>");
    officeXMLFile.write("</Row>");

    // ====[ Tom rad mellan veckoavsnitten ]====
    officeXMLFile.write("<Row>");
    officeXMLFile.write("<Cell ss:Index=\"2\" ss:StyleID=\"s82\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s82\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s82\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s82\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s82\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s82\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s82\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s82\"/>");
    officeXMLFile.write("<Cell ss:StyleID=\"s82\"/>");
    officeXMLFile.write("</Row>");

  }

  // ======================================================================================================================================
  //   Skriv ut data för en dag 
  // ======================================================================================================================================
  private static void WriteDayData(int vecka, int dag) throws IOException {
    
    // ====[ Om det är söndag eller lördag - rödmarkera raden och skriv HELG som kommentar ]====
    if(cal.get(Calendar.DAY_OF_WEEK)!=1 && cal.get(Calendar.DAY_OF_WEEK)!=7) {
      
      officeXMLFile.write("<Row>");
      officeXMLFile.write("<Cell ss:Index=\"2\" ss:StyleID=\"veckonr\" ss:Formula=\"=WEEKNUM(RC[1],21)\">"); // Formula för att beräkna veckonummer, hämtar datumvärdet
      officeXMLFile.write("<Data ss:Type=\"Number\" />");
      officeXMLFile.write("</Cell>");
      
      
      // ====[ Den först cellen [C3] måste ha ett riktigt datum, efterföljande hämtar ifrån denna cellen+1 ]=====
      if(vecka==1 & dag==1) {
        
        officeXMLFile.write("<Cell ss:StyleID=\"s66\">");
        officeXMLFile.write("<Data ss:Type=\"DateTime\">");
        
        // ====[ Skriv ut datumet för den första cellen, resterande kommer beräknas i Excel med hjälp av formulas ]====
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS");
        officeXMLFile.write(sdf.format(cal.getTime()));
        officeXMLFile.write("</Data>");
        
      } else {
        
        // ====[ Hämta värdet ifrån föregående veckoavsnitt (söndagen) om det är en ny vecka ]====
        if(cal.get(Calendar.DAY_OF_WEEK)==2) {
          officeXMLFile.write("<Cell ss:StyleID=\"s66\" ss:Formula=\"=R[-4]C+1\">");  
        } else {
          officeXMLFile.write("<Cell ss:StyleID=\"s66\" ss:Formula=\"=R[-1]C+1\">");
        }
        
        officeXMLFile.write("<Data ss:Type=\"DateTime\" />");

      }

      officeXMLFile.write("</Cell>");
      officeXMLFile.write("<Cell ss:StyleID=\"s66\" ss:Formula=\"=TEXT(RC[-1],&quot;ddd&quot;)\">"); // Formulan för att beräkna vilken dag det är, exempelvis sön, mån utifrån datumvärdet. Språkberoende
      officeXMLFile.write("<Data ss:Type=\"String\" />");
      officeXMLFile.write("</Cell>");
      officeXMLFile.write("<Cell ss:StyleID=\"tidkol\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"tidkol\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"tidkol\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"tidkol\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"tidkol\" ss:Formula=\"=(RC[-3]-RC[-4])-(RC[-1]-RC[-2])\">"); // Beräkna antal timmar och dra bort lunchen
      officeXMLFile.write("<Data ss:Type=\"DateTime\" />");
      officeXMLFile.write("</Cell>");
      officeXMLFile.write("<Cell ss:StyleID=\"s68\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"s69\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"s70\"/>");
      officeXMLFile.write("</Row>");
    } else {
      
      // ====[ H E L G ]====
      officeXMLFile.write("<Row>");
      officeXMLFile.write("<Cell ss:Index=\"2\" ss:StyleID=\"s71\" ss:Formula=\"=WEEKNUM(RC[1],21)\">"); // Formula för att beräkna veckonummer, hämtar datumvärdet
      officeXMLFile.write("<Data ss:Type=\"Number\" />");
      officeXMLFile.write("</Cell>");
      
      // ====[ Den först cellen måste ha ett riktigt datum, efterföljande hämtar ifrån denna cellen+1 ]=====
      if(vecka==1 & dag==1) {
        
        officeXMLFile.write("<!-- SPECIAL -->");
        officeXMLFile.write("<Cell ss:StyleID=\"s72\">");
        officeXMLFile.write("<Data ss:Type=\"DateTime\">");
        
        // ====[ Skriv ut datumet för den första cellen, resterande kommer beräknas i Excel med hjälp av formulas ]====
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS");
        officeXMLFile.write(sdf.format(cal.getTime()));
        officeXMLFile.write("</Data>");
        
      } else {
        
        officeXMLFile.write("<!-- VANLIG -->");
        // ====[ Hämta värdet ifrån föregående veckoavsnitt (söndagen) om det är en ny vecka ]====
        if(cal.get(Calendar.DAY_OF_WEEK)==2) {
          officeXMLFile.write("<Cell ss:StyleID=\"s72\" ss:Formula=\"=R[-4]C+1\">");  
        } else {
          officeXMLFile.write("<Cell ss:StyleID=\"s72\" ss:Formula=\"=R[-1]C+1\">");
        }
        
        officeXMLFile.write("<Data ss:Type=\"DateTime\" />");

      }

      officeXMLFile.write("</Cell>");
      officeXMLFile.write("<Cell ss:StyleID=\"s72\" ss:Formula=\"=TEXT(RC[-1],&quot;ddd&quot;)\">"); // Formulan för att beräkna vilken dag det är, exempelvis sön, mån utifrån datumvärdet. Språkberoende
      officeXMLFile.write("<Data ss:Type=\"String\" />");
      officeXMLFile.write("</Cell>");
      officeXMLFile.write("<Cell ss:StyleID=\"s73\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"s73\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"s73\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"s73\"/>");
      officeXMLFile.write("<Cell ss:StyleID=\"s73\" ss:Formula=\"=(RC[-3]-RC[-4])-(RC[-1]-RC[-2])\">"); // Beräkna antal timmar och dra bort lunchen
      officeXMLFile.write("<Data ss:Type=\"DateTime\" />");
      officeXMLFile.write("</Cell>");
      officeXMLFile.write("<Cell ss:StyleID=\"s74\"/>");
      officeXMLFile.write("<Cell ss:Index=\"12\" ss:StyleID=\"s75\">");
      officeXMLFile.write("<Data ss:Type=\"String\">  H E L G</Data>"); // Skriv ut helg i anteckningarna
      officeXMLFile.write("</Cell>");
      officeXMLFile.write("</Row>");
    }    
  }

  // ======================================================================================================================================
  //   Skriv ut huvudet som finns ovanför varje avsnitt
  // ======================================================================================================================================
  private static void WriteTableHeader(int veckonr) throws IOException {
    
    if(veckonr==1) {
      officeXMLFile.write("<Row ss:Index=\"2\">");  // Börja en rad ned i Kalkylbladet så det blir lite luft i överkant
    } else {
      officeXMLFile.write("<Row>");  
    }
    
    officeXMLFile.write("<Cell ss:Index=\"2\" ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Veckonr</Data></Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Datum</Data></Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Veckodag</Data></Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Kom</Data></Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Gick</Data></Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Lunch</Data></Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Lunch</Data></Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Tid</Data></Cell>");
    officeXMLFile.write("<Cell ss:StyleID=\"huvud\"><Data ss:Type=\"String\">Tid</Data></Cell>");
    officeXMLFile.write("<Cell ss:Index=\"12\" ss:StyleID=\"s64\"><Data ss:Type=\"String\">  Anteckningar mm</Data></Cell>");
    officeXMLFile.write("</Row>");
  }

  // ======================================================================================================================================
  //  Inställningar för Office dokument
  // ======================================================================================================================================
  private static void WriteOfficeDocumentSettings() throws IOException {
    officeXMLFile.write("<OfficeDocumentSettings xmlns=\"urn:schemas-microsoft-com:office:office\"><AllowPNG/></OfficeDocumentSettings>");
  }

  // ======================================================================================================================================
  //  Arbetsboken
  // ======================================================================================================================================
  private static void WriteExcelWorkbook() throws IOException {
    officeXMLFile.write("<ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\"><WindowHeight>12300</WindowHeight><WindowWidth>28800</WindowWidth><WindowTopX>0</WindowTopX><WindowTopY>0</WindowTopY><ProtectStructure>False</ProtectStructure><ProtectWindows>False</ProtectWindows></ExcelWorkbook>");
  }

  //===================================================================
  //  Skriv ut koden för formatteringen, fonter, färger linjer mm
  //===================================================================
  private static void WriteStyles() throws IOException {
    
    officeXMLFile.write("<Styles>");
    officeXMLFile.write("<Style ss:ID=\"Default\" ss:Name=\"Normal\"><Alignment ss:Vertical=\"Bottom\"/><Borders/><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/><Interior/><NumberFormat/><Protection/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s62\"><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/></Style>");
    officeXMLFile.write("<!-- ====[ Svart rubrik med vit text ]==== --><Style ss:ID=\"huvud\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#000000\" ss:Pattern=\"Solid\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s64\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#FFFFFF\"/><Interior ss:Color=\"#000000\" ss:Pattern=\"Solid\"/></Style>");
    officeXMLFile.write("<!-- ====[ Veckonummer ]==== --><Style ss:ID=\"veckonr\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s66\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior/><NumberFormat ss:Format=\"Short Date\"/></Style>");
    officeXMLFile.write("<!-- ====[ Tidskolumnen ]==== --><Style ss:ID=\"tidkol\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior/><NumberFormat ss:Format=\"Short Time\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s68\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior/><NumberFormat ss:Format=\"Fixed\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s69\"><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s70\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s71\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior ss:Color=\"#FF0000\" ss:Pattern=\"Solid\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s72\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior ss:Color=\"#FF0000\" ss:Pattern=\"Solid\"/><NumberFormat ss:Format=\"Short Date\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s73\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior ss:Color=\"#FF0000\" ss:Pattern=\"Solid\"/><NumberFormat ss:Format=\"Short Time\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s74\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior ss:Color=\"#FF0000\" ss:Pattern=\"Solid\"/><NumberFormat ss:Format=\"Fixed\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s75\"><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><Interior ss:Color=\"#FF0000\" ss:Pattern=\"Solid\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s76\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s77\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/><NumberFormat ss:Format=\"Short Date\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s78\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s79\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\" ss:Bold=\"1\" ss:Italic=\"1\"/><Interior ss:Color=\"#FFFF00\" ss:Pattern=\"Solid\"/><NumberFormat ss:Format=\"[h]:mm:ss;@\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s80\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\" ss:Bold=\"1\" ss:Italic=\"1\"/><Interior ss:Color=\"#FFFF00\" ss:Pattern=\"Solid\"/><NumberFormat ss:Format=\"Fixed\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s81\"><Borders><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/></Style>");
    officeXMLFile.write("<Style ss:ID=\"s82\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Top\"/><Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Color=\"#000000\"/></Style>");
    officeXMLFile.write("</Styles>");
    
  }

  // ======================================================================================================================================
  //    Skapar avsnittet DocumentProperties ( Dokumentegenskaperna )
  // ======================================================================================================================================
  private static void WriteDocumentProperties() throws IOException {

    SimpleDateFormat skapandeDatum = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'");
    Date datum = new Date();

    StringBuffer kod = new StringBuffer();

    // ====[ Fyll i lite data i dokumentegenskaperna för arbetsboken ]====
    kod.append("<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">");
    kod.append("<Title>En kalender i Excel</Title>");
    kod.append("<Author>Java Magic</Author>");
    kod.append("<Keywords>Kalender</Keywords>");
    kod.append("<LastAuthor>Java Magic</LastAuthor>");
    kod.append("<Created>").append(skapandeDatum.format(datum)).append("</Created>\n");
    kod.append("<LastSaved>").append(skapandeDatum.format(datum)).append("</LastSaved>");
    kod.append("<Category>Tidrapportering</Category>");
    kod.append("<Version>14.00</Version>").append("</DocumentProperties>");

    officeXMLFile.write(kod.toString());

  }

  // ======================================================================================================================================
  //   Skriv ut XML deklarationerna och deklarationen för Arbetsboken
  // ======================================================================================================================================
  private static void WriteFileStart() throws IOException {

    // ====[ XML deklarationer ]====
    officeXMLFile.write("<?xml version=\"1.0\"?><?mso-application progid=\"Excel.Sheet\"?>");
    officeXMLFile.write("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">");
  }

  // ======================================================================================================================================
  //   Stäng Arbetsboken
  // ======================================================================================================================================
  private static void WriteFileEnd() throws IOException {
    officeXMLFile.write("</Worksheet></Workbook>");
  }

  // ======================================================================================================================================
  //  Kalkylbladsinställningar
  // ======================================================================================================================================
  private static void WriteWorkSheetOptions() throws IOException {
    officeXMLFile.write("<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\"><PageSetup><Header x:Margin=\"0.3\"/><Footer x:Margin=\"0.3\"/><PageMargins x:Bottom=\"0.75\" x:Left=\"0.7\" x:Right=\"0.7\" x:Top=\"0.75\"/></PageSetup><Print><ValidPrinterInfo/><PaperSizeIndex>9</PaperSizeIndex><HorizontalResolution>600</HorizontalResolution><VerticalResolution>600</VerticalResolution></Print><Selected/><DoNotDisplayGridlines/><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions>");
  }

  // ======================================================================================================================================
  //   Kontrollera om ett värde är numeriskt
  // ======================================================================================================================================
  public static boolean isNumeric(String strNum) {
    
    if (strNum == null) {
      return false;
    }
    
    try {
      int x = Integer.parseInt(strNum);
    } catch (NumberFormatException nfe) {
      return false;
    }
    return true;
  }
}
