package fdi.ucm.server.exportparser.xls.gramcount;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map.Entry;

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
import fdi.ucm.server.modelComplete.collection.grammar.CompleteOperationalValueType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteResourceElementType;

public class CollectionXLSGramCount {

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

        HashMap<CompleteGrammar, List<CompleteElementType>> GramarElem=new HashMap<>();
        
        for (CompleteGrammar gram : salvar.getMetamodelGrammar()) {
        	if (!isIgnore(gram))
        	{
			List<CompleteElementType> ListaGramar=new LinkedList<>();
			GramarElem.put(gram, ListaGramar);
			processList(gram.getSons(),ListaGramar);
        	}
		}
        
        int ImagenesCount=0;
        
        for (Entry<CompleteGrammar, List<CompleteElementType>> entryGram : GramarElem.entrySet()) {
        	ImagenesCount=ImagenesCount+procesaHoja(entryGram.getKey().getNombre(),entryGram.getValue(),libro,cL,salvar.getEstructuras());
		}
        
        System.out.println(ImagenesCount);
 
        libro.write(archivo);

        archivo.close();

		return rutaArchivo;

        
	}

	
	
	private static int procesaHoja(String nombre, List<CompleteElementType> listValidos, Workbook libro,
			CompleteLogAndUpdates cL, List<CompleteDocuments> documentos) {
		int Salida = 0;

		Sheet hoja;
		hoja = libro.createSheet(nombre);
		int row = 0;
		Row filaH = hoja.createRow(row);
		row++;
		Cell celda0 = filaH.createCell(0);
		celda0.setCellValue(nombre+" ID");
		Cell celda1 = filaH.createCell(1);
		celda1.setCellValue(nombre+" Desc.");
		Cell celda2 = filaH.createCell(2);
		celda2.setCellValue("#Elements");
		
		
		for (CompleteDocuments completeDocuments : documentos) {

			int ContadorValidosEnGramatica = 0;
			for (CompleteElement elemento : completeDocuments.getDescription())
				if (listValidos.contains(elemento.getHastype()))
					ContadorValidosEnGramatica++;
			
			for (CompleteElement elemento : completeDocuments.getDescription())
				if (listValidos.contains(elemento.getHastype())&&
						elemento.getHastype() instanceof CompleteResourceElementType &&
						!isNoImage(elemento.getHastype()))
					Salida++;

			if (ContadorValidosEnGramatica>0)
			{
				Row fila = hoja.createRow(row);
				row++;
				Cell celda01 = fila.createCell(0);
				celda01.setCellValue(completeDocuments.getClavilenoid());
				Cell celda11 = filaH.createCell(1);
				celda11.setCellValue(completeDocuments.getDescriptionText());
				Cell celda21 = filaH.createCell(2);
				celda21.setCellValue(ContadorValidosEnGramatica);		
			}

		}

		return Salida;

	}



	private static boolean isIgnore(CompleteElementType hastype) {
		for (CompleteOperationalValueType ovT : hastype.getShows()) {
			if (ovT.getView().toLowerCase().equals("xml")
				&&ovT.getName().toLowerCase().equals("ignore"))
				try {
					return Boolean.parseBoolean(ovT.getDefault());
				} catch (Exception e) {
					// Sigo a mi rollo a ver si hay otro
				}
		}
		return false;
	}



	private static boolean isNoImage(CompleteElementType hastype) {
		for (CompleteOperationalValueType ovT : hastype.getShows()) {
			if (ovT.getView().toLowerCase().equals("xml")
				&&ovT.getName().toLowerCase().equals("noimage"))
				try {
					return Boolean.parseBoolean(ovT.getDefault());
				} catch (Exception e) {
					// Sigo a mi rollo a ver si hay otro
				}
		}
		return false;
	}



	private static boolean isIgnore(CompleteGrammar gram) {
		for (CompleteOperationalValueType ovT : gram.getViews()) {
			if (ovT.getView().toLowerCase().equals("xml")
				&&ovT.getName().toLowerCase().equals("ignore"))
				try {
					return Boolean.parseBoolean(ovT.getDefault());
				} catch (Exception e) {
					// Sigo a mi rollo a ver si hay otro
				}
		}
		return false;
	}





	private static void processList(List<CompleteElementType> sons, List<CompleteElementType> ordenTerms) {
		for (CompleteElementType term : sons) {
			if (!isIgnore(term)) 
			ordenTerms.add(term);
			processList(term.getSons(),ordenTerms);
		}
		
	}





public static void main(String[] args) throws Exception{
	
		
		String message="Exception .clavy-> Params Null ";
		try {

			
			
			String fileName = "test.clavy";
			
			if (args.length!=0)
				fileName=args[1];
			
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
