
import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import javax.swing.JTextArea;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import java.io.FileOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;









public class reworkoneday{

	public void writeoneday(String datedebut , String datefin , int valuetosend , String valuemed1 , String valuemed2 , String valuemed3 , String valuemed4 , String valuemed5 , JTextArea loadingarea) {  
		
		
		
		String newdatedebut = "" ; 
		String newdatefin = "" ;
		
		
		if(datedebut.equals("")) {
			 newdatedebut = "";
		}else {
			 String slashh = "/" ; 
			 String anneedatedebut = datedebut.substring(0,4);
		     String moisdatedebut = datedebut.substring(4,6);
		     String jourdatedebut = datedebut.substring(6);
		     newdatedebut =  jourdatedebut + slashh + moisdatedebut + slashh + anneedatedebut  ;
		}
 
	     if(datefin.equals("")) {
	    	  newdatefin = "" ;
	     }else {
			 String slashh = "/" ; 
	    	 String anneedatefin = datefin.substring(0,4);
		     String moisdatefin = datefin.substring(4,6);
		     String jourdatefin = datefin.substring(6);
		      newdatefin =  jourdatefin + slashh + moisdatefin + slashh + anneedatefin  ;
	     }
	     
	    
	
	
	
	
	
	File wordFile1 = new File("./Sortie.doc");
	
    
    wordFile1.delete();

    
    
    XWPFDocument document = new XWPFDocument();
    XWPFParagraph title = document.createParagraph();
    XWPFParagraph paragraph = document.createParagraph();
    
    XWPFRun titre = title.createRun();
	titre.setText("ONEDAY");
	titre.addBreak();
	titre.setText("DU " + newdatedebut + " AU " + newdatefin);
	titre.addBreak();
	titre.addBreak(BreakType.PAGE);
	titre.setFontSize(35);
	titre.setFontFamily("ARIAL BLACK");
	titre.setBold(true);
	title.setAlignment(ParagraphAlignment.CENTER);
	
        
		
		String query = "";
		String pickspenum = "" ;
		String newligne=System.getProperty("line.separator");
		String ligne = ".................................................................................................................................";
		String lignepoint = "___________________________________________________________________________";

		
		
			
		
    	String newentryhr  = ""; 
    	String newentrydt  = ""; 
    	String newopdt = "" ;
    	String newophr = "" ;
    	String slash = "" ; 
        String anneeadmi = "";
        String moisadmi = "";
        String jouradmi = "";
        String heureadmi = "";
        String minutesadmi = "";
        String secondesadmi = "";
        String anneeop = "";
        String moisop = "";
        String jourop = "";
        String point = "" ; 
        String heureop = "";
        String minutesop = "";
        String secondesop = "";
        String sortie = "";
        String premedname = "";
        String predateop = "" ;

		
		
		ArrayList<String> spenum = new ArrayList<String>();
	       
    	spenum.add("0");
    	spenum.add("140");
    	spenum.add("210");
    	spenum.add("240");
    	spenum.add("310");
    	spenum.add("340");
    	spenum.add("370");
    	spenum.add("410");
    	spenum.add("450");
    	spenum.add("480");
    	spenum.add("520");
    	spenum.add("550");
    	spenum.add("650");
    	spenum.add("A12");
    	spenum.add("ASS");
    	spenum.add("MED");
    	spenum.add("URG");
    	spenum.add("IM");
    	spenum.add("MAT");
    	spenum.add("001");
    	spenum.add("003");
    	spenum.add("014");
    	spenum.add("034");
    	spenum.add("048");
    	spenum.add("100");
    	spenum.add("170");
    	spenum.add("414");
    	spenum.add("580");
    	spenum.add("620");
    	spenum.add("690");
    	spenum.add("000");
    	spenum.add("010");
    	spenum.add("270");
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
    	
		if (datefin != "") {
			
			


				try {
					Class.forName("");
					Connection con = DriverManager.getConnection("");
					
				    	
				    	pickspenum = spenum.get(valuetosend);
						System.out.println(pickspenum);

						//AUCUN MEDECIN
						if (valuemed1.equals("") & valuemed2.equals("") & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{
				    	if(pickspenum.equals("0")) {
							System.out.println("Aucun medecin");
							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";				
				    	}
				    	else {				    		
							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.ANESNAME , OP.MEDSPE , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";
				    	}
						}
						
						//1 MEDECIN 
						else if (valuemed1 != "" & valuemed2.equals("") & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{				
					    	if(pickspenum.equals("0")) {	
								System.out.println("1 medecin");

						    	query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" + datedebut +"' AND '" + datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB = '" + valuemed1 + "' ) AND ( OP.SURGNB = '" + valuemed1 + "' ) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(OPDT)";	
					    	}
					    	else {
					    	query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" + datedebut +"' AND '" + datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB = '" + valuemed1 + "' ) AND ( OP.SURGNB = '" + valuemed1 + "' ) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(OPDT)";	
					    	}
						}
						
						//2 MEDECIN
						else if (valuemed1 != "" & valuemed2 != "" & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{	
						if(pickspenum.equals("0")) {				    		
							System.out.println("2 medecin");

							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" 
							+ datedebut +"' AND '" 
							+ datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"')) AND MEDBLOC.MEDNB = OP.SURGNB   ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME";													    		
				    	}
				    	else {			
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" 
							+ datedebut +"' AND '" 
							+ datefin + "' ) AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"')) AND MEDBLOC.MEDNB = OP.SURGNB   ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME";				
				    	}
						}
						
						//3 MEDECIN
						else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4.equals("") & valuemed5.equals("") )	{
				    	if(pickspenum.equals("0")) {			    		
							System.out.println("3 medecin");

				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" 
							+ datedebut +"' AND '" + datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";											    		
				    	}
				    	else {				    		
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" 
							+ datedebut +"' AND '" + datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";											
				    	}
						}
						
						//4 MEDECIN
						else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4 != "" & valuemed5.equals("") )	{						
						if(pickspenum.equals("0")) {
							System.out.println("4 medecin");

							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" 
							+ datedebut +"' AND '" + datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"
							+valuemed4+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";												    		
				    	}
				    	else {			    		
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX ,  OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" 
							+ datedebut +"' AND '" + datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"
							+valuemed4+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";		
				    	}
						}
				    	
				    	//5 MEDECIN
						else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4 != "" & valuemed5 != "" )	{
						if(pickspenum.equals("0")) {
							System.out.println("5 medecin");

							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX ,  OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" 
							+ datedebut +"' AND '" + datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND (OP.SURGNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";
				    	}
				    	else{
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT BETWEEN '" 
							+ datedebut +"' AND '" + datefin + "' ) AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND (OP.SURGNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";
				    	}
						}
				    	
				    	
				    	
						System.out.println(query);
						Statement stmt=con.createStatement();  
						ResultSet rs = stmt.executeQuery(query);
							
					
						
						int compteurpage = 0 ; 
						
						
						for (int i = 0 ; i < 500; i++) {
							
							
							compteurpage++ ;
							
							
							rs.next();
							String medfirstname = rs.getString("MEDFIRSTNAME");
							String medname = rs.getString("MEDNAME");
							String patnb = rs.getString("PATNB");
							String patname = rs.getString("PATNAME");
							String patfirstname = rs.getString("PATFIRSTNAME");
							String entrydt = rs.getString("ENTRYDT");
							String patsex = rs.getString("PATSEX");
							String heureopdebut = rs.getString("OPEXPBEGHR");
							String litdem = rs.getString("LITDEM");
							String roomid = rs.getString("ROOMID"); 
							String entryhr = rs.getString("ENTRYHR");
							String dateop = rs.getString("OPDT") ; 
							String typeanes = rs.getString("ANESTYP");
							String nomop = rs.getString("OPCOMNAME");
							String anesname = rs.getString("ANESNAME");
							String latextest = rs.getString("LATEX");

							
							
							XWPFRun nommed = paragraph.createRun();
							XWPFRun date = paragraph.createRun();
							XWPFRun nom = paragraph.createRun();
							XWPFRun nomresult = paragraph.createRun();
							XWPFRun prenom = paragraph.createRun();
							XWPFRun prenomresult = paragraph.createRun();
							XWPFRun sexe = paragraph.createRun();
							XWPFRun sexeresult = paragraph.createRun();
							XWPFRun du = paragraph.createRun();
							XWPFRun duresult = paragraph.createRun();
							XWPFRun datehradmi = paragraph.createRun();
							XWPFRun datehradmiresult = paragraph.createRun();
							XWPFRun datehrop = paragraph.createRun();
							XWPFRun datehropresult = paragraph.createRun();
							XWPFRun opname = paragraph.createRun();
							XWPFRun opnameresult = paragraph.createRun();
							XWPFRun anest = paragraph.createRun();
							XWPFRun anestresult= paragraph.createRun();
							XWPFRun anestype= paragraph.createRun();
							XWPFRun anestyperesult = paragraph.createRun();
							XWPFRun salleop = paragraph.createRun();
							XWPFRun salleopresult = paragraph.createRun();
							XWPFRun litask = paragraph.createRun();
							XWPFRun litaskresult  = paragraph.createRun();
							
							paragraph.setAlignment(ParagraphAlignment.LEFT);
			            
							nommed.setFontSize(12);
							nom.setFontSize(9);
							nomresult.setFontSize(9);
							prenom.setFontSize(9);
							prenomresult.setFontSize(9);
							anest.setFontSize(9);
							anestresult.setFontSize(9);
							du.setFontSize(9);
							duresult.setFontSize(9);
							sexe.setFontSize(9);
							sexeresult.setFontSize(9);
							datehradmi.setFontSize(9);
							datehradmiresult.setFontSize(9);
							datehrop.setFontSize(9);
							datehropresult.setFontSize(9);
							opname.setFontSize(9);
							opnameresult.setFontSize(9);
							anestype.setFontSize(9);
							anestyperesult.setFontSize(9);
							salleop.setFontSize(9);
							salleopresult.setFontSize(9);
							litask.setFontSize(9);
							litaskresult.setFontSize(9);
			            
			            
							nommed.setBold(true);
							nom.setBold(true);
							nomresult.setBold(true);
							prenom.setBold(true);
							prenomresult.setBold(true);
			          	 	du.setBold(true);
			            
			            
							
							
							
			          	 	
							 if (entryhr.equals("")) {
								 	
								 entrydt = " " ;
								 
							 } else {
								 
								 slash = "/" ; 
							     anneeadmi = entrydt.substring(0,4);
							     moisadmi = entrydt.substring(4,6);
							     jouradmi = entrydt.substring(6);
							     newentrydt =  jouradmi + slash + moisadmi + slash + anneeadmi  ;
							 } 

							 if (dateop.equals("")) {
								  
								 dateop = "" ;
								 
							 } else {
								 
								 slash = "/" ; 
								 anneeop = dateop.substring(0,4);
							     moisop = dateop.substring(4,6);
							     jourop = dateop.substring(6);
							     newopdt =  jourop + slash + moisop + slash + anneeop  ;
								 
							 }
							 
			            
							if (entryhr.equals("")) {
								
								newentryhr = "" ;
								
							}
							
							else {
								 point = ":" ; 
							     heureadmi = entryhr.substring(0,2);
							     minutesadmi = entryhr.substring(2,4);
							     secondesadmi = entryhr.substring(4);
							 
							     newentryhr =  heureadmi + point + minutesadmi + point + secondesadmi  ;
			         
							}
							
							
							if (heureopdebut.equals("")) {
								
								newophr = "" ;

								
							} else {
								 point = ":" ; 
							     heureop = heureopdebut.substring(0,2);
							     minutesop = heureopdebut.substring(2,4);
							     secondesop = heureopdebut.substring(4);
							    
							    
							    newophr =  heureop + point + minutesop + point + secondesop  ;
								
							}

			           
							
							if(compteurpage % 5 == 0) {
								 
								
								if(dateop.equals(predateop)){
									
									if (medname.equals(premedname)) {
										
										
										nommed.addBreak(BreakType.PAGE);
										nommed.addBreak();
										nommed.setText(lignepoint);
										nommed.addBreak();
										nommed.setText("Dr . " + medname);
										date.setText("       Date : " + newopdt);
										date.addBreak();
										date.setText(lignepoint);
										date.addBreak();
										date.addBreak();
										nom.setText("Nom : ");
										nomresult.setText("  " + patname);
										prenom.setText("      Pr�nom :  ");
										prenomresult.setText(patfirstname);
										sexe.setText("           Sexe : " + patsex);
										sexe.addBreak();
										du.setText("DU :  " + patnb);
										du.addBreak();
										datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
										datehradmi.addBreak();
										datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
										datehrop.addBreak();
										opname.setText("Intervention : " + nomop);
										anest.setText("  Anesth : Dr.  " + anesname);
										anest.addBreak();
										anestype.setText( "Anesth�sie  : " + typeanes );
										salleop.setText("   Salle OP : " + roomid);
										litask.setText("   Lit : " + litdem);
										litask.addBreak();
										if(latextest.equals("O")) {
											
											litask.setText("ALLERGIE AU LATEX");

										}
										
										litask.addBreak();
										litask.setText(ligne);
										litask.addBreak();


										
										
										

									}else {
										
										nommed.addBreak(BreakType.PAGE);
										nommed.addBreak();
										nommed.setText(lignepoint);
										nommed.addBreak();
										nommed.setText("Dr . " + medname);
										date.setText("       Date : " + newopdt);
										date.addBreak();
										date.setText(lignepoint);
										date.addBreak();
										date.addBreak();
											nom.setText("Nom : ");
											nomresult.setText("  " + patname);
											prenom.setText("      Pr�nom :  ");
											prenomresult.setText(patfirstname);
											sexe.setText("    Sexe : " + patsex);
											sexe.addBreak();
											du.setText("DU :  " + patnb);
											du.addBreak();
											datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
											datehradmi.addBreak();
											datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
											datehrop.addBreak();
											opname.setText("Intervention : " + nomop);
											anest.setText("  Anesth : Dr.  " + anesname);
											anest.addBreak();
											anestype.setText( "Anesth�sie  : " + typeanes );
											salleop.setText("   Salle OP : " + roomid);
											litask.setText("   Lit : " + litdem);
											litask.addBreak();
											if(latextest.equals("O")) {
												
												litask.setText("ALLERGIE AU LATEX");

											}
											
											litask.addBreak();
											litask.setText(ligne);
											litask.addBreak();



											


					            }
					            
									
									
									
									
								}else {
									
									if (medname.equals(premedname)) {
										
										nommed.addBreak(BreakType.PAGE);
										nommed.addBreak();
										nommed.setText(lignepoint);
										nommed.addBreak();
										nommed.setText("Dr . " + medname);
										date.setText("       Date : " + newopdt);
										date.addBreak();
										date.setText(lignepoint);
										date.addBreak();
										date.addBreak();
										nom.setText("Nom : ");
										nomresult.setText("  " + patname);
										prenom.setText("      Pr�nom :  ");
										prenomresult.setText(patfirstname);
										sexe.setText("           Sexe : " + patsex);
										sexe.addBreak();
										du.setText("DU :  " + patnb);
										du.addBreak();
										datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
										datehradmi.addBreak();
										datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
										datehrop.addBreak();
										opname.setText("Intervention : " + nomop);
										anest.setText("  Anesth : Dr.  " + anesname);
										anest.addBreak();
										anestype.setText( "Anesth�sie  : " + typeanes );
										salleop.setText("   Salle OP : " + roomid);
										litask.setText("   Lit : " + litdem);
										litask.addBreak();
										if(latextest.equals("O")) {
											
											litask.setText("ALLERGIE AU LATEX");

										}
										
										litask.addBreak();
										litask.setText(ligne);
										litask.addBreak();


										
										
										

									}else {
										
											nommed.addBreak(BreakType.PAGE);
											nommed.addBreak();
											nommed.setText(lignepoint);
											nommed.addBreak();
											nommed.setText("Dr . " + medname);
											date.setText("       Date : " + newopdt);
											date.addBreak();
											date.setText(lignepoint);
											date.addBreak();
											date.addBreak();
											
											nom.setText("Nom : ");
											nomresult.setText("  " + patname);
											prenom.setText("      Pr�nom :  ");
											prenomresult.setText(patfirstname);
											sexe.setText("    Sexe : " + patsex);
											sexe.addBreak();
											du.setText("DU :  " + patnb);
											du.addBreak();
											datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
											datehradmi.addBreak();
											datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
											datehrop.addBreak();
											opname.setText("Intervention : " + nomop);
											anest.setText("  Anesth : Dr.  " + anesname);
											anest.addBreak();
											anestype.setText( "Anesth�sie  : " + typeanes );
											salleop.setText("   Salle OP : " + roomid);
											litask.setText("   Lit : " + litdem);
											litask.addBreak();
											if(latextest.equals("O")) {
												
												litask.setText("ALLERGIE AU LATEX");

											}
											
											litask.addBreak();
											litask.setText(ligne);
											litask.addBreak();



											


					            }
					            
									
									
									
									
								}
				             
								
								
								
							}else {
								
								
								
								
								if(dateop.equals(predateop)){
									
									

									
									if (medname.equals(premedname)) {
										
										
										
										nom.setText("Nom : ");
										nomresult.setText("  " + patname);
										prenom.setText("      Pr�nom :  ");
										prenomresult.setText(patfirstname);
										sexe.setText("           Sexe : " + patsex);
										sexe.addBreak();
										du.setText("DU :  " + patnb);
										du.addBreak();
										datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
										datehradmi.addBreak();
										datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
										datehrop.addBreak();
										opname.setText("Intervention : " + nomop);
										anest.setText("  Anesth : Dr.  " + anesname);
										anest.addBreak();
										anestype.setText( "Anesth�sie  : " + typeanes );
										salleop.setText("   Salle OP : " + roomid);
										litask.setText("   Lit : " + litdem);
										litask.addBreak();
										if(latextest.equals("O")) {
											
											litask.setText("ALLERGIE AU LATEX");

										}
										
										litask.addBreak();
										litask.setText(ligne);
										litask.addBreak();


										
										
										

									}
									
									

									else {
										
											nommed.addBreak();
											nommed.setText(lignepoint);
											nommed.addBreak();
											nommed.setText("Dr . " + medname);
											date.setText("       Date : " + newopdt);
											date.addBreak();
											date.setText(lignepoint);
											date.addBreak();
											date.addBreak();
											
											nom.setText("Nom : ");
											nomresult.setText("  " + patname);
											prenom.setText("      Pr�nom :  ");
											prenomresult.setText(patfirstname);
											sexe.setText("    Sexe : " + patsex);
											sexe.addBreak();
											du.setText("DU :  " + patnb);
											du.addBreak();
											datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
											datehradmi.addBreak();
											datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
											datehrop.addBreak();
											opname.setText("Intervention : " + nomop);
											anest.setText("  Anesth : Dr.  " + anesname);
											anest.addBreak();
											anestype.setText( "Anesth�sie  : " + typeanes );
											salleop.setText("   Salle OP : " + roomid);
											litask.setText("   Lit : " + litdem);
											litask.addBreak();
											if(latextest.equals("O")) {
												
												litask.setText("ALLERGIE AU LATEX");

											}
											
											litask.addBreak();
											litask.setText(ligne);
											litask.addBreak();



											


					            }
					            
									
									
									
									
								}
								
								//FAMDIFDT
								
								else {
									
									
									
									if (medname.equals(premedname)) {

										compteurpage = 0 ; 
										
										
										nommed.addBreak(BreakType.PAGE);
										nommed.addBreak();
										nommed.setText(lignepoint);
										nommed.addBreak();
										nommed.setText("Dr . " + medname);
										date.setText("       Date : " + newopdt);
										date.addBreak();
										date.setText(lignepoint);
										date.addBreak();
										date.addBreak();
										nom.setText("Nom : ");
										nomresult.setText("  " + patname);
										prenom.setText("      Pr�nom :  ");
										prenomresult.setText(patfirstname);
										sexe.setText("           Sexe : " + patsex);
										sexe.addBreak();
										du.setText("DU :  " + patnb);
										du.addBreak();
										datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
										datehradmi.addBreak();
										datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
										datehrop.addBreak();
										opname.setText("Intervention : " + nomop);
										anest.setText("  Anesth : Dr.  " + anesname);
										anest.addBreak();
										anestype.setText( "Anesth�sie  : " + typeanes );
										salleop.setText("   Salle OP : " + roomid);
										litask.setText("   Lit : " + litdem);
										litask.addBreak();
										if(latextest.equals("O")) {
											
											litask.setText("ALLERGIE AU LATEX");

										}
										
										litask.addBreak();
										litask.setText(ligne);
										litask.addBreak();



										
										
										

									}
									

									else {
										
										
										
											compteurpage = 0 ; 
										
											
											nommed.addBreak(BreakType.PAGE);
											nommed.addBreak();
											nommed.setText(lignepoint);
											nommed.addBreak();
											nommed.setText("Dr . " + medname);
											date.setText("       Date : " + newopdt);
											date.addBreak();
											date.setText(lignepoint);
											date.addBreak();
											date.addBreak();
											
											nom.setText("Nom : ");
											nomresult.setText("  " + patname);
											prenom.setText("      Pr�nom :  ");
											prenomresult.setText(patfirstname);
											sexe.setText("    Sexe : " + patsex);
											sexe.addBreak();
											du.setText("DU :  " + patnb);
											du.addBreak();
											datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
											datehradmi.addBreak();
											datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
											datehrop.addBreak();
											opname.setText("Intervention : " + nomop);
											anest.setText("  Anesth : Dr.  " + anesname);
											anest.addBreak();
											anestype.setText( "Anesth�sie  : " + typeanes );
											salleop.setText("   Salle OP : " + roomid);
											litask.setText("   Lit : " + litdem);
											litask.addBreak();
											if(latextest.equals("O")) {
												
												litask.setText("ALLERGIE AU LATEX");

											}
											
											litask.addBreak();
											litask.setText(ligne);
											litask.addBreak();


					            }
					            
		
								}
				        	
							}
			
									premedname = medname ; 
									predateop = dateop ;


						      
						 
						    
						}	
					} catch(Exception e){ System.out.println(e);}
					
					
							try {
								 FileOutputStream output = new FileOutputStream("Sortie.doc") ;
								document.write(output);
								output.close();
								
							}catch(Exception e ) {
								e.printStackTrace();
							}
							
							
							
							
					loadingarea.setText("Cr�ation du fichier termin� ! ");
				
	
		}
		
		
		// PAS DE DATE DE FIN ( JOUR ACTUEL )

		else if (datefin == "") {
			

				System.out.println("Aucun medecin");

				try {
					Class.forName("");
					Connection con = DriverManager.getConnection("");
					
				    	
				    	pickspenum = spenum.get(valuetosend);
						

				    	
				    	//AUCUN MEDECIN
						if (valuemed1.equals("") & valuemed2.equals("") & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{
				    	if(pickspenum.equals("0")) {
							System.out.println("Aucun medecin");
							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";				
				    	}
				    	else {				    		
							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.ANESNAME , OP.MEDSPE , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";
				    	}
						}
						
						//1 MEDECIN 
						else if (valuemed1 != "" & valuemed2.equals("") & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{				
					    	if(pickspenum.equals("0")) {	
								System.out.println("1 medecin");

						    	query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB = '" + valuemed1 + "' ) AND ( OP.SURGNB = '" + valuemed1 + "' ) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(OPDT)";	
					    	}
					    	else {
					    	query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB = '" + valuemed1 + "' ) AND ( OP.SURGNB = '" + valuemed1 + "' ) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(OPDT)";	
					    	}
						}
						
						//2 MEDECIN
						else if (valuemed1 != "" & valuemed2 != "" & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{	
						if(pickspenum.equals("0")) {				    		
							System.out.println("2 medecin");

							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"')) AND MEDBLOC.MEDNB = OP.SURGNB   ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME";													    		
				    	}
				    	else {			
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"')) AND MEDBLOC.MEDNB = OP.SURGNB   ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME";				
				    	}
						}
						
						//3 MEDECIN
						else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4.equals("") & valuemed5.equals("") )	{
				    	if(pickspenum.equals("0")) {			    		
							System.out.println("3 medecin");

				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";											    		
				    	}
				    	else {				    		
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";											
				    	}
						}
						
						//4 MEDECIN
						else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4 != "" & valuemed5.equals("") )	{						
						if(pickspenum.equals("0")) {
							System.out.println("4 medecin");

							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"
							+valuemed4+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";												    		
				    	}
				    	else {			    		
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX ,  OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"
							+valuemed4+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";		
				    	}
						}
				    	
				    	//5 MEDECIN
						else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4 != "" & valuemed5 != "" )	{
						if(pickspenum.equals("0")) {
							System.out.println("5 medecin");

							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX ,  OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND (OP.SURGNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";
				    	}
				    	else{
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.SURGNB , OP.LATEX , OP.PATNB , OP.ENTRYDT , OP.MEDSPE , OP.ANESNAME , OP.OPEXPBEGHR , OP.LITDEM , OP.ROOMID , OP.ENTRYHR , OP.OPDT , OP.ANESTYP , OP.OPCOMNAME , MEDBLOC.MEDNB , MEDBLOC.MEDNAME , MEDBLOC.MEDFIRSTNAME FROM OP , MEDBLOC WHERE (OP.OPDT LIKE '" + datedebut + "') AND OP.TYPHOSP LIKE 'O' AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND (OP.SURGNB IN ('"
							+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(OPDT) , MEDBLOC.MEDNAME ";
				    	}
						}
				    	
						
						System.out.println(query);
						Statement stmt=con.createStatement();  
						ResultSet rs = stmt.executeQuery(query);
						
					
						int compteurpage = 0 ; 
						
						
						for (int i = 0 ; i < 500; i++) {
							
							
							compteurpage++ ;
							System.out.println(compteurpage);
							
							
							rs.next();
							String medfirstname = rs.getString("MEDFIRSTNAME");
							String medname = rs.getString("MEDNAME");
							String patnb = rs.getString("PATNB");
							String patname = rs.getString("PATNAME");
							String patfirstname = rs.getString("PATFIRSTNAME");
							String entrydt = rs.getString("ENTRYDT");
							String patsex = rs.getString("PATSEX");
							String heureopdebut = rs.getString("OPEXPBEGHR");
							String litdem = rs.getString("LITDEM");
							String roomid = rs.getString("ROOMID"); 
							String entryhr = rs.getString("ENTRYHR");
							String dateop = rs.getString("OPDT") ; 
							String typeanes = rs.getString("ANESTYP");
							String nomop = rs.getString("OPCOMNAME");
							String anesname = rs.getString("ANESNAME");
							
							
							//NAMXWPF

							
							XWPFRun nommed = paragraph.createRun();
							XWPFRun date = paragraph.createRun();
							XWPFRun nom = paragraph.createRun();
							XWPFRun nomresult = paragraph.createRun();
							XWPFRun prenom = paragraph.createRun();
							XWPFRun prenomresult = paragraph.createRun();
							XWPFRun sexe = paragraph.createRun();
							XWPFRun sexeresult = paragraph.createRun();
							XWPFRun du = paragraph.createRun();
							XWPFRun duresult = paragraph.createRun();
							XWPFRun datehradmi = paragraph.createRun();
							XWPFRun datehradmiresult = paragraph.createRun();
							XWPFRun datehrop = paragraph.createRun();
							XWPFRun datehropresult = paragraph.createRun();
							XWPFRun opname = paragraph.createRun();
							XWPFRun opnameresult = paragraph.createRun();
							XWPFRun anest = paragraph.createRun();
							XWPFRun anestresult= paragraph.createRun();
							XWPFRun anestype= paragraph.createRun();
							XWPFRun anestyperesult = paragraph.createRun();
							XWPFRun salleop = paragraph.createRun();
							XWPFRun salleopresult = paragraph.createRun();
							XWPFRun litask = paragraph.createRun();
							XWPFRun litaskresult  = paragraph.createRun();
							
							paragraph.setAlignment(ParagraphAlignment.LEFT);
			            
			            //TAILLE 
							nommed.setFontSize(12);
							nom.setFontSize(9);
							nomresult.setFontSize(9);
							prenom.setFontSize(9);
							prenomresult.setFontSize(9);
							anest.setFontSize(9);
							anestresult.setFontSize(9);
							du.setFontSize(9);
							duresult.setFontSize(9);
							sexe.setFontSize(9);
							sexeresult.setFontSize(9);
							datehradmi.setFontSize(9);
							datehradmiresult.setFontSize(9);
							datehrop.setFontSize(9);
							datehropresult.setFontSize(9);
							opname.setFontSize(9);
							opnameresult.setFontSize(9);
							anestype.setFontSize(9);
							anestyperesult.setFontSize(9);
							salleop.setFontSize(9);
							salleopresult.setFontSize(9);
							litask.setFontSize(9);
							litaskresult.setFontSize(9);
			            
			            
							//POLICE
							nommed.setBold(true);
							nom.setBold(true);
							nomresult.setBold(true);
							prenom.setBold(true);
							prenomresult.setBold(true);
			          	 	du.setBold(true);
			            
			            
							
							
							//NAMDTHR
							
			          	 	
							 if (entryhr.equals("")) {
								 	
								 entrydt = " " ;
								 
							 } else {
								 
								 slash = "/" ; 
							     anneeadmi = entrydt.substring(0,4);
							     moisadmi = entrydt.substring(4,6);
							     jouradmi = entrydt.substring(6);
							     newentrydt =  jouradmi + slash + moisadmi + slash + anneeadmi  ;
							 } 

							 if (dateop.equals("")) {
								  
								 dateop = "" ;
								 
							 } else {
								 
								 slash = "/" ; 
								 anneeop = dateop.substring(0,4);
							     moisop = dateop.substring(4,6);
							     jourop = dateop.substring(6);
							     newopdt =  jourop + slash + moisop + slash + anneeop  ;
								 
							 }
							 
			            
							if (entryhr.equals("")) {
								
								newentryhr = "" ;
								
							}
							
							else {
								 point = ":" ; 
							     heureadmi = entryhr.substring(0,2);
							     minutesadmi = entryhr.substring(2,4);
							     secondesadmi = entryhr.substring(4);
							 
							     newentryhr =  heureadmi + point + minutesadmi + point + secondesadmi  ;
			         
							}
							
							
							if (heureopdebut.equals("")) {
								
								newophr = "" ;

								
							} else {
								 point = ":" ; 
							     heureop = heureopdebut.substring(0,2);
							     minutesop = heureopdebut.substring(2,4);
							     secondesop = heureopdebut.substring(4);
							    
							    
							    newophr =  heureop + point + minutesop + point + secondesop  ;
								
							}

							
							//NCOMPTJUMP0

							if(compteurpage % 5 == 0) {
								 
								
									
									if (medname.equals(premedname)) {
										
										
										nommed.addBreak(BreakType.PAGE);
										nommed.addBreak();
										nommed.setText(lignepoint);
										nommed.addBreak();
										nommed.setText("Dr . " + medname);
										date.setText("       Date : " + newopdt);
										date.addBreak();
										date.setText(lignepoint);
										date.addBreak();
										date.addBreak();
										nom.setText("Nom : ");
										nomresult.setText("  " + patname);
										prenom.setText("      Pr�nom :  ");
										prenomresult.setText(patfirstname);
										sexe.setText("           Sexe : " + patsex);
										sexe.addBreak();
										du.setText("DU :  " + patnb);
										du.addBreak();
										datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
										datehradmi.addBreak();
										datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
										datehrop.addBreak();
										opname.setText("Intervention : " + nomop);
										anest.setText("  Anesth : Dr.  " + anesname);
										anest.addBreak();
										anestype.setText( "Anesth�sie  : " + typeanes );
										salleop.setText("   Salle OP : " + roomid);
										litask.setText("   Lit : " + litdem);
										litask.addBreak();
										litask.setText(ligne);
										litask.addBreak();
			

									}else {
										
										nommed.addBreak(BreakType.PAGE);
										nommed.addBreak();
										nommed.setText(lignepoint);
										nommed.addBreak();
										nommed.setText("Dr . " + medname);
										date.setText("       Date : " + newopdt);
										date.addBreak();
										date.setText(lignepoint);
										date.addBreak();
										date.addBreak();
											nom.setText("Nom : ");
											nomresult.setText("  " + patname);
											prenom.setText("      Pr�nom :  ");
											prenomresult.setText(patfirstname);
											sexe.setText("    Sexe : " + patsex);
											sexe.addBreak();
											du.setText("DU :  " + patnb);
											du.addBreak();
											datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
											datehradmi.addBreak();
											datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
											datehrop.addBreak();
											opname.setText("Intervention : " + nomop);
											anest.setText("  Anesth : Dr.  " + anesname);
											anest.addBreak();
											anestype.setText( "Anesth�sie  : " + typeanes );
											salleop.setText("   Salle OP : " + roomid);
											litask.setText("   Lit : " + litdem);
											litask.addBreak();
											litask.setText(ligne);
											litask.addBreak();

					            }
					            
				
							}else {
								
									//NAMSAMEMED


									if (medname.equals(premedname)) {
										
										
										
										nom.setText("Nom : ");
										nomresult.setText("  " + patname);
										prenom.setText("      Pr�nom :  ");
										prenomresult.setText(patfirstname);
										sexe.setText("           Sexe : " + patsex);
										sexe.addBreak();
										du.setText("DU :  " + patnb);
										du.addBreak();
										datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
										datehradmi.addBreak();
										datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
										datehrop.addBreak();
										opname.setText("Intervention : " + nomop);
										anest.setText("  Anesth : Dr.  " + anesname);
										anest.addBreak();
										anestype.setText( "Anesth�sie  : " + typeanes );
										salleop.setText("   Salle OP : " + roomid);
										litask.setText("   Lit : " + litdem);
										litask.addBreak();
										litask.setText(ligne);
										litask.addBreak();


										
										
										

									}
									
									//NAMDIFMED


									else {
										
											nommed.addBreak();
											nommed.setText(lignepoint);
											nommed.addBreak();
											nommed.setText("Dr . " + medname);
											date.setText("       Date : " + newopdt);
											date.addBreak();
											date.setText(lignepoint);
											date.addBreak();
											date.addBreak();
											
											nom.setText("Nom : ");
											nomresult.setText("  " + patname);
											prenom.setText("      Pr�nom :  ");
											prenomresult.setText(patfirstname);
											sexe.setText("    Sexe : " + patsex);
											sexe.addBreak();
											du.setText("DU :  " + patnb);
											du.addBreak();
											datehradmi.setText("Date et heure d'admission     : "+ newentrydt  + " � " + newentryhr);
											datehradmi.addBreak();
											datehrop.setText("Date et heure d'intervention : " + newopdt + " � " + newophr);
											datehrop.addBreak();
											opname.setText("Intervention : " + nomop);
											anest.setText("  Anesth : Dr.  " + anesname);
											anest.addBreak();
											anestype.setText( "Anesth�sie  : " + typeanes );
											salleop.setText("   Salle OP : " + roomid);
											litask.setText("   Lit : " + litdem);
											litask.addBreak();
											litask.setText(ligne);
											litask.addBreak();



											


					            }
					            
									
									
									
									
								}
								
								
							
							 
							
			             
							
			          
									

									premedname = medname ; 
									predateop = dateop ;


						      
						 
						    
						}	
					} catch(Exception e){ System.out.println(e);}
					
					
							try {
								 FileOutputStream output = new FileOutputStream("Sortie.doc") ;
								document.write(output);
								output.close();
								
							}catch(Exception e ) {
								e.printStackTrace();
							}
							
							
							
							
							
					loadingarea.setText("Cr�ation du fichier termin� ! ");
				
			
			
			
			
			
		}
	}

	}

