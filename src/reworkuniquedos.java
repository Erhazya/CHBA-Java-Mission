
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
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;









public class reworkuniquedos{

	public void writeuniquedos(String datedebut , String datefin , int valuetosend , String valuemed1 , String valuemed2 , String valuemed3 , String valuemed4 , String valuemed5 , JTextArea loadingarea) {  
		
		
		
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

        
        
        try (XWPFDocument document = new XWPFDocument()) {
			XWPFParagraph title = document.createParagraph();
			XWPFParagraph paragraph = document.createParagraph();
			
			XWPFRun titre = title.createRun();
			titre.setText("DEMANDE DE DOSSIERS");
			titre.addBreak();
			titre.setText("DU " + newdatedebut + " AU " + newdatefin);
			titre.addBreak();
			titre.addBreak(BreakType.PAGE);
			titre.addBreak(BreakType.PAGE);
			titre.setFontSize(35);
			titre.setFontFamily("ARIAL BLACK");
			titre.setBold(true);
			title.setAlignment(ParagraphAlignment.CENTER);


			

			
			String query = "";
			String pickspenum = "" ;
			String ligne = ".................................................................................................................................";
			String lignepoint = "___________________________________________________________________________";

			
			
				
			
			String newentrydt  = ""; 
			String slash = "" ; 
			String anneeop = "";
			String moisop = "";
			String jourop = "";
			String predateentry = "" ;
			String anneenaiss = "";
			String moisnaiss = "";
			String journaiss = "";
			String newdtnaiss = "";
			
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
						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '"+ datedebut +"' AND '" + datefin + "' ) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) , MEDBLOC.MEDNAME ";				
			    	}
			    	else {				    		
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '"+ datedebut +"' AND '" + datefin + "' ) AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) , MEDBLOC.MEDNAME ";
			    	}
					}
					
					//1 MEDECIN 
					else if (valuemed1 != "" & valuemed2.equals("") & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{				
				    	if(pickspenum.equals("0")) {	
							System.out.println("1 medecin");

							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" + datedebut +"' AND '" + datefin + "' ) AND (MEDBLOC.MEDNB = '" + valuemed1 + "' ) AND ( OP.SURGNB = '" + valuemed1 + "' ) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(ENTRYDT)";	
				    	}
				    	else {
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" + datedebut +"' AND '" + datefin + "' ) AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB = '" + valuemed1 + "' ) AND ( OP.SURGNB = '" + valuemed1 + "' ) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(ENTRYDT)";	
				    	}
					}
					
					//2 MEDECIN
					else if (valuemed1 != "" & valuemed2 != "" & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{	
					if(pickspenum.equals("0")) {				    		
						System.out.println("2 medecin");

						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" 
						+ datedebut +"' AND '" 
						+ datefin + "' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"')) AND MEDBLOC.MEDNB = OP.SURGNB   ORDER BY ABS(ENTRYDT)";													    		
			    	}
			    	else {			
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" 
						+ datedebut +"' AND '" 
						+ datefin + "' ) AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"')) AND MEDBLOC.MEDNB = OP.SURGNB   ORDER BY ABS(ENTRYDT)";				
			    	}
					}
					
					//3 MEDECIN
					else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4.equals("") & valuemed5.equals("") )	{
			    	if(pickspenum.equals("0")) {			    		
						System.out.println("3 medecin");

						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" 
						+ datedebut +"' AND '" + datefin + "' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) ";											    		
			    	}
			    	else {				    		
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" 
						+ datedebut +"' AND '" + datefin + "' ) AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) ";											
			    	}
					}
					
					//4 MEDECIN
					else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4 != "" & valuemed5.equals("") )	{						
					if(pickspenum.equals("0")) {
						System.out.println("4 medecin");

						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" 
						+ datedebut +"' AND '" + datefin + "' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"
						+valuemed4+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) ";												    		
			    	}
			    	else {			    		
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" 
						+ datedebut +"' AND '" + datefin + "' ) AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"
						+valuemed4+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT)  ";		
			    	}
					}
			    	
			    	//5 MEDECIN
					else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4 != "" & valuemed5 != "" )	{
					if(pickspenum.equals("0")) {
						System.out.println("5 medecin");

						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" 
						+ datedebut +"' AND '" + datefin + "' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND (OP.SURGNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(ENTRYDT) ";
			    	}
			    	else{
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB FROM OP , MEDBLOC WHERE (OP.ENTRYDT BETWEEN '" 
						+ datedebut +"' AND '" + datefin + "' ) AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND (OP.SURGNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(ENTRYDT) ";
			    	}
					}
			    	
					
					
						System.out.println(query);
						Statement stmt=con.createStatement();  
						ResultSet rs = stmt.executeQuery(query);
							
						System.out.println("1");

						
						int compteurpage = 0 ; 
						
						
						for (int i = 0 ; i < 500; i++) {

							
							compteurpage++ ;
							System.out.println(compteurpage);

							
							
							rs.next();
							String patnb = rs.getString("PATNB");
							String patname = rs.getString("PATNAME");
							String patfirstname = rs.getString("PATFIRSTNAME");
							String patsex = rs.getString("PATSEX");
							String dateentry = rs.getString("ENTRYDT");
							String naissdt = rs.getString("PATBIRTHDAY");
															  

							
							//FAMXWPF
							
							XWPFRun date = paragraph.createRun();
							XWPFRun nom = paragraph.createRun();
							XWPFRun nomresult = paragraph.createRun();
							XWPFRun prenom = paragraph.createRun();
							XWPFRun prenomresult = paragraph.createRun();
							XWPFRun sexe = paragraph.createRun();
							XWPFRun sexeresult = paragraph.createRun();
							XWPFRun datenaiss = paragraph.createRun();
							XWPFRun du = paragraph.createRun();
							XWPFRun duresult = paragraph.createRun();
							
							paragraph.setAlignment(ParagraphAlignment.LEFT);
			            
			            //TAILLE 
							nom.setFontSize(9);
							nomresult.setFontSize(9);
							prenom.setFontSize(9);
							prenomresult.setFontSize(9);
							du.setFontSize(9);
							duresult.setFontSize(9);
							sexe.setFontSize(9);
							sexeresult.setFontSize(9);
							datenaiss.setFontSize(9);

			            
							//POLICE
							nom.setBold(true);
							prenom.setBold(true);
			          	 	du.setBold(true);
			            
			            
							
							
							
			          	 	//FAMDTHR
			          	 	
							 

							 if (dateentry.equals("")) {
								  
								 dateentry = "" ;
								 
							 } else {
								 
								 slash = "/" ; 
								 anneeop = dateentry.substring(0,4);
							     moisop = dateentry.substring(4,6);
							     jourop = dateentry.substring(6);
							     newentrydt =  jourop + slash + moisop + slash + anneeop  ;
								 
							 }
							 
			            
							 if (naissdt.equals("")) {
								  
								 naissdt = "" ;
								 
							 } else {
								 
								 slash = "/" ; 
								 anneenaiss = naissdt.substring(0,4);
							     moisnaiss = naissdt.substring(4,6);
							     journaiss = naissdt.substring(6);
							     newdtnaiss =  journaiss + slash + moisnaiss + slash + anneenaiss  ;
								 
							 }
							 
							
							
							

			           
							//FCOMPTJUMP0
							
							if(compteurpage % 15 == 0) {
								 								
										
										date.addBreak(BreakType.PAGE);
										date.setText(lignepoint);
										date.addBreak();
										date.setText("Date : " + newentrydt);
										date.addBreak();
										date.setText(lignepoint);
										date.addBreak();
										nom.setText("Nom : ");
										nomresult.setText("  " + patname);
										prenom.setText("      Pr�nom :  ");
										prenomresult.setText(patfirstname);
										sexe.setText("           Sexe : " + patsex);
										datenaiss.setText("        Date Naiss : " + newdtnaiss);
										datenaiss.addBreak();
										du.setText("DU :  " + patnb);
										du.addBreak();
										

			
									
}
			
			else {
								
								
								
					if(dateentry.equals(predateentry)){
						
						
					
						nom.addBreak();
						nom.setText("Nom : ");
						nomresult.setText("  " + patname);
						prenom.setText("      Pr�nom :  ");
						prenomresult.setText(patfirstname);
						sexe.setText("           Sexe : " + patsex);
						datenaiss.setText("        Date Naiss : " + newdtnaiss);
						datenaiss.addBreak();
						du.setText("DU :  " + patnb);
						du.addBreak();


					
					
					

				}else {
					
					date.addBreak();
					date.setText(lignepoint);
					date.addBreak();
					date.setText("Date : " + newentrydt);
					date.addBreak();
					date.setText(lignepoint);
					date.addBreak();
					nom.setText("Nom : ");
					nomresult.setText("  " + patname);
					prenom.setText("      Pr�nom :  ");
					prenomresult.setText(patfirstname);
					sexe.setText("    Sexe : " + patsex);
					datenaiss.setText("        Date Naiss : " + newdtnaiss);
					datenaiss.addBreak();
					du.setText("DU :  " + patnb);
					du.addBreak();


						


}
						
				             
								
								
							}
							 
							
			             
							
			          
									

									predateentry = dateentry ;


						      
						 
						    
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
			
			
			
			
				
			
			else if (datefin == "") {
				
				
				try {
					Class.forName("");
					Connection con = DriverManager.getConnection("");
					
				    	


				    				    		
					pickspenum = spenum.get(valuetosend);
					System.out.println(pickspenum);


					//AUCUN MEDECIN
					if (valuemed1.equals("") & valuemed2.equals("") & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{
			    	if(pickspenum.equals("0")) {
						System.out.println("Aucun medecin");
						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) , MEDBLOC.MEDNAME ";				
			    	}
			    	else {				    		
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) , MEDBLOC.MEDNAME ";
			    	}
					}
					
					//1 MEDECIN 
					else if (valuemed1 != "" & valuemed2.equals("") & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{				
				    	if(pickspenum.equals("0")) {	
							System.out.println("1 medecin");

							query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (MEDBLOC.MEDNB = '" + valuemed1 + "' ) AND ( OP.SURGNB = '" + valuemed1 + "' ) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(ENTRYDT)";	
				    	}
				    	else {
				    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB = '" + valuemed1 + "' ) AND ( OP.SURGNB = '" + valuemed1 + "' ) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(ENTRYDT)";	
				    	}
					}
					
					//2 MEDECIN
					else if (valuemed1 != "" & valuemed2 != "" & valuemed3.equals("") & valuemed4.equals("") & valuemed5.equals("") )	{	
					if(pickspenum.equals("0")) {				    		
						System.out.println("2 medecin");

						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"')) AND MEDBLOC.MEDNB = OP.SURGNB   ORDER BY ABS(ENTRYDT)";													    		
			    	}
			    	else {			
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"')) AND MEDBLOC.MEDNB = OP.SURGNB   ORDER BY ABS(ENTRYDT)";				
			    	}
					}
					
					//3 MEDECIN
					else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4.equals("") & valuemed5.equals("") )	{
			    	if(pickspenum.equals("0")) {			    		
						System.out.println("3 medecin");

						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) ";											    		
			    	}
			    	else {				    		
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) ";											
			    	}
					}
					
					//4 MEDECIN
					else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4 != "" & valuemed5.equals("") )	{						
					if(pickspenum.equals("0")) {
						System.out.println("4 medecin");

						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"
						+valuemed4+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT) ";												    		
			    	}
			    	else {			    		
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"')) AND (OP.SURGNB IN ('"+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"
						+valuemed4+"')) AND MEDBLOC.MEDNB = OP.SURGNB ORDER BY ABS(ENTRYDT)  ";		
			    	}
					}
			    	
			    	//5 MEDECIN
					else if (valuemed1 != "" & valuemed2 != "" & valuemed3 != "" & valuemed4 != "" & valuemed5 != "" )	{
					if(pickspenum.equals("0")) {
						System.out.println("5 medecin");

						query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB  FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND (OP.SURGNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(ENTRYDT) ";
			    	}
			    	else{
			    		query = "SELECT OP.PATNAME , OP.PATFIRSTNAME, OP.PATSEX , OP.PATBIRTHDAY , OP.PATNB , OP.ENTRYDT , OP.SURGNB , MEDBLOC.MEDNB FROM OP , MEDBLOC WHERE (OP.ENTRYDT LIKE '" + datedebut + "') AND (OP.MEDSPE LIKE '" + pickspenum +"' ) AND (MEDBLOC.MEDNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND (OP.SURGNB IN ('"
						+valuemed1+"','"+valuemed2+"','"+valuemed3+"','"+valuemed4+"','"+valuemed5+"')) AND MEDBLOC.MEDNB = OP.SURGNB  ORDER BY ABS(ENTRYDT) ";
			    	}
					}
			    	

						
						
						System.out.println(query);
						Statement stmt=con.createStatement();  
						ResultSet rs = stmt.executeQuery(query);
							
					
						
						int compteurpage = 0 ; 
						
						
						for (int i = 0 ; i < 500; i++) {
							
							
							compteurpage++ ;
							
							
							rs.next();
							String patnb = rs.getString("PATNB");
							String patname = rs.getString("PATNAME");
							String patfirstname = rs.getString("PATFIRSTNAME");
							String patsex = rs.getString("PATSEX");
							String dateentry = rs.getString("ENTRYDT");
							String naissdt = rs.getString("PATBIRTHDAY");


							
							//FAMXWPF
							
							XWPFRun date = paragraph.createRun();
							XWPFRun nom = paragraph.createRun();
							XWPFRun nomresult = paragraph.createRun();
							XWPFRun prenom = paragraph.createRun();
							XWPFRun prenomresult = paragraph.createRun();
							XWPFRun sexe = paragraph.createRun();
							XWPFRun sexeresult = paragraph.createRun();
							XWPFRun datenaiss = paragraph.createRun();
							XWPFRun du = paragraph.createRun();
							XWPFRun duresult = paragraph.createRun();
							
							paragraph.setAlignment(ParagraphAlignment.LEFT);
			            
			            //TAILLE 
							nom.setFontSize(9);
							nomresult.setFontSize(9);
							prenom.setFontSize(9);
							prenomresult.setFontSize(9);
							du.setFontSize(9);
							duresult.setFontSize(9);
							sexe.setFontSize(9);
							sexeresult.setFontSize(9);
							datenaiss.setFontSize(9);

			            
							//POLICE
							nom.setBold(true);
							nomresult.setBold(true);
							prenom.setBold(true);
							prenomresult.setBold(true);
			          	 	du.setBold(true);
			            
			            
							
							
							
			          	 	//FAMDTHR
			          	 	
							 

							 if (dateentry.equals("")) {
								  
								 dateentry = "" ;
								 
							 } else {
								 
								 slash = "/" ; 
								 anneeop = dateentry.substring(0,4);
							     moisop = dateentry.substring(4,6);
							     jourop = dateentry.substring(6);
							     newentrydt =  jourop + slash + moisop + slash + anneeop  ;
								 
							 }
							 
							 if (naissdt.equals("")) {
								  
								 naissdt = "" ;
								 
							 } else {
								 
								 slash = "/" ; 
								 anneenaiss = naissdt.substring(0,4);
							     moisnaiss = naissdt.substring(4,6);
							     journaiss = naissdt.substring(6);
							     newdtnaiss =  journaiss + slash + moisnaiss + slash + anneenaiss  ;
								 
							 }
							 
							
							
							 
							

			           
							//FCOMPTJUMP0
							
							if(compteurpage % 15 == 0) {
								 								
										
										date.addBreak(BreakType.PAGE);
										date.addBreak();
										date.setText("Date : " + newentrydt);
										date.addBreak();
										date.setText(lignepoint);
										date.addBreak();
										nom.setText("Nom : ");
										nomresult.setText("  " + patname);
										prenom.setText("      Pr�nom :  ");
										prenomresult.setText(patfirstname);
										sexe.setText("           Sexe : " + patsex);
										datenaiss.setText("        Date Naiss : " + newdtnaiss);
										datenaiss.addBreak();
										du.setText("DU :  " + patnb);
										du.addBreak();
										du.setText(ligne);
										

			
									
}
			
			else {
								
								
								
					if(dateentry.equals(predateentry)){
						
						
					
						nom.addBreak();
						nom.setText("Nom : ");
						nomresult.setText("  " + patname);
						prenom.setText("      Pr�nom :  ");
						prenomresult.setText(patfirstname);
						sexe.setText("           Sexe : " + patsex);
						datenaiss.setText("        Date Naiss : " + newdtnaiss);
						datenaiss.addBreak();
						du.setText("DU :  " + patnb);
						du.addBreak();
						du.setText(ligne);


					
					
					

				}else {
					
					date.addBreak();
					date.setText(lignepoint);
					date.addBreak();
					date.setText("Date : " + newentrydt);
					date.addBreak();
					date.setText(lignepoint);
					date.addBreak();
					date.addBreak();
					nom.setText("Nom : ");
					nomresult.setText("  " + patname);
					prenom.setText("      Pr�nom :  ");
					prenomresult.setText(patfirstname);
					sexe.setText("    Sexe : " + patsex);
					datenaiss.setText("        Date Naiss : " + newdtnaiss);
					datenaiss.addBreak();
					du.setText("DU :  " + patnb);
					du.addBreak();
					du.setText(ligne);


						


}
						
				             
								
								
							}
							 
							
			             
							
			          
									

									predateentry = dateentry ;


						      
						 
						    
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
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
			
			 
			
		
			
		}
		
		

		
		
	}
	

	

