package fdi.ucm.server.exportparser.xls.termcount;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.CompleteLogAndUpdates;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;

public class CollectionXLSTermCount {

	public static String processCompleteCollection(CompleteLogAndUpdates cL, CompleteCollection salvar,
			boolean soloEstructura, String pathTemporalFiles) throws IOException{
		 /*La ruta donde se creará el archivo*/
        String rutaArchivo = pathTemporalFiles+"/"+System.nanoTime()+".xls";
        /*Se crea el objeto de tipo File con la ruta del archivo*/
        File archivoXLS = new File(rutaArchivo);
        /*Si el archivo existe se elimina*/
        if(archivoXLS.exists()) archivoXLS.delete();
        /*Se crea el archivo*/
        archivoXLS.createNewFile();
        
        /*Se crea el libro de excel usando el objeto de tipo Workbook*/
        Workbook libro = new HSSFWorkbook();
        
        /*Se inicializa el flujo de datos con el archivo xls*/
        FileOutputStream archivo = new FileOutputStream(archivoXLS);

        
//        Sheet hoja;
        processHojaDocumentos(libro,salvar.getEstructuras(),cL);
        processHojaTopicos(libro,salvar.getEstructuras(),salvar.getMetamodelGrammar(),cL);
      
        libro.write(archivo);

        archivo.close();

		return rutaArchivo;

        
	}

	
	
	private static void processHojaTopicos(Workbook libro, List<CompleteDocuments> estructuras,
			List<CompleteGrammar> metamodelGrammar, CompleteLogAndUpdates cL) {
		
		HashMap<Long, Integer> Id_NElems=new HashMap<>();
		List<CompleteElementType> OrdenTerms=new LinkedList<CompleteElementType>();
		
		for (CompleteGrammar completeGrammar : metamodelGrammar) 
			processList(completeGrammar.getSons(),OrdenTerms);
		
		for (CompleteDocuments doc : estructuras) {
			for (CompleteElement elem : doc.getDescription()) {
				CompleteElementType elemType = elem.getHastype().getClassOfIterator();
				if (elemType==null)
					elemType=elem.getHastype();
				
				if (elemType!=null)
				{
					Integer Cantidad=Id_NElems.get(elemType.getClavilenoid());
					if (Cantidad==null)
						Cantidad=0;
					
					Cantidad=new Integer(Cantidad.intValue()+1);
					Id_NElems.put(elemType.getClavilenoid(), Cantidad);
					
				}
			}
		}
		
		Sheet hoja;
		hoja = libro.createSheet("Concepts");
		int row = 0;
		Row filaH = hoja.createRow(row);
		row++;
		Cell celda0 = filaH.createCell(0);
		celda0.setCellValue("Concept ID");
		Cell celda1 = filaH.createCell(1);
		celda1.setCellValue("Concept Name.");
		Cell celda2 = filaH.createCell(2);
		celda2.setCellValue("#Elements");
		for (CompleteElementType elem : OrdenTerms) {
			Integer Cantidad=Id_NElems.get(elem.getClavilenoid());
			if (Cantidad==null)
				Cantidad=0;
			Row fila = hoja.createRow(row);
			row++;
			Cell celda01 = fila.createCell(0);
			celda01.setCellValue(elem.getClavilenoid());
			Cell celda11 = filaH.createCell(1);
			celda11.setCellValue(elem.getName());
			Cell celda21 = filaH.createCell(2);
			celda21.setCellValue(Cantidad);
		}
		
		
	}



	private static void processList(List<CompleteElementType> sons, List<CompleteElementType> ordenTerms) {
		for (CompleteElementType term : sons) {
			ordenTerms.add(term);
			processList(term.getSons(),ordenTerms);
		}
		
	}



	private static void processHojaDocumentos(Workbook libro, List<CompleteDocuments> estructuras,
			CompleteLogAndUpdates cL) {
		Sheet hoja;
		hoja = libro.createSheet("Documents");
		int row = 0;
		Row filaH = hoja.createRow(row);
		row++;
		Cell celda0 = filaH.createCell(0);
		celda0.setCellValue("Document ID");
		Cell celda1 = filaH.createCell(1);
		celda1.setCellValue("Document Desc.");
		Cell celda2 = filaH.createCell(2);
		celda2.setCellValue("#Elements");

		for (CompleteDocuments docU : estructuras) {
			Row fila = hoja.createRow(row);
			row++;
			Cell celda01 = fila.createCell(0);
			celda01.setCellValue(docU.getClavilenoid());
			Cell celda11 = filaH.createCell(1);
			celda11.setCellValue(docU.getDescriptionText());
			Cell celda21 = filaH.createCell(2);
			celda21.setCellValue(docU.getDescription().size());

		}
	}



public static void main(String[] args) throws Exception{
	
		
		String message="Exception .clavy-> Params Null ";
		try {

			
			
			String fileName = "test.clavy";
			 System.out.println(fileName);
			 

			 File file = new File(fileName);
			 FileInputStream fis = new FileInputStream(file);
			 ObjectInputStream ois = new ObjectInputStream(fis);
			 CompleteCollection object = (CompleteCollection) ois.readObject();
			 
			 
			 try {
				 ois.close();
			} catch (Exception e) {
				// TODO: handle exception
			}
			
			 try {
				 fis.close();
			} catch (Exception e) {
				// TODO: handle exception
			}
			 
			 
		
		 
		  
		  
			 processCompleteCollection(new CompleteLogAndUpdates(), object, false, System.getProperty("user.home"));
			 
	    }catch (Exception e) {
			e.printStackTrace();
			System.err.println(message);
			throw new RuntimeException(message);
		}
		  
		  
		 
		  
	    }
	
}
