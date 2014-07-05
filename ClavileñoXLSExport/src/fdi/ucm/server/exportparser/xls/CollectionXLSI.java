/**
 * 
 */
package fdi.ucm.server.exportparser.xls;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.CompleteCollectionLog;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteLinkElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteResourceElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteStructure;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;

/**
 * @author Joaquin Gayoso-Cabada
 *Clase qie produce el XLSI
 */
public class CollectionXLSI {


	public static String processCompleteCollection(CompleteCollectionLog cL,
			CompleteCollection salvar, boolean soloEstructura, String pathTemporalFiles) throws IOException {
		
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
        
        /*Utilizamos la clase Sheet para crear una nueva hoja de trabajo dentro del libro que creamos anteriormente*/
        
        
        Sheet hoja;
        
        if (!salvar.getName().isEmpty())
        	 hoja = libro.createSheet(salvar.getName());
        else hoja = libro.createSheet();
        
        
        
       
        
        ArrayList<CompleteElementType> ListaElementos=generaLista(salvar.getMetamodelGrammar());
        
        
        /*Hacemos un ciclo para inicializar los valores de 10 filas de celdas*/
        for(int f=0;f<(salvar.getEstructuras().size()+1);f++){
            /*La clase Row nos permitirá crear las filas*/
            Row fila = hoja.createRow(f);

            /*Cada fila tendrá 5 celdas de datos*/
            for(int c=0;c<ListaElementos.size();c++){
                /*Creamos la celda a partir de la fila actual*/
                Cell celda = fila.createCell(c);
                
                /*Si la fila es la número 0, estableceremos los encabezados*/
                if(f==0){
                    celda.setCellValue(ListaElementos.get(c).getName());
                }else{
                    /*Si no es la primera fila establecemos un valor*/
                    celda.setCellValue("Valor celda "+c+","+f);
                }
            }
        }
        
        /*Escribimos en el libro*/
        libro.write(archivo);
        /*Cerramos el flujo de datos*/
        archivo.close();
        /*Y abrimos el archivo con la clase Desktop*/
////        Desktop.getDesktop().open(archivoXLS);
		return rutaArchivo;
	}
	
	  private static ArrayList<CompleteElementType> generaLista(
			List<CompleteGrammar> metamodelGrammar) {
		  ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		  for (CompleteGrammar completegramar : metamodelGrammar) {
			ListaElementos.addAll(generaLista(completegramar));
		}
		return ListaElementos;
	}

	private static Collection<CompleteElementType> generaLista(
			CompleteGrammar completegramar) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 for (CompleteStructure completeelem : completegramar.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType)
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
				ListaElementos.addAll(generaLista(completeelem));
			}
		 return ListaElementos;
	}

	private static Collection<? extends CompleteElementType> generaLista(
			CompleteStructure completeelementPadre) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 for (CompleteStructure completeelem : completeelementPadre.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType)
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
				ListaElementos.addAll(generaLista(completeelem));
			}
		 return ListaElementos;
	}

	public static void main(String[] args) throws Exception{
		  
		  CompleteCollection CC=new CompleteCollection("Lou Arreglate", "Arreglate ya!");
		  for (int i = 0; i < 10; i++) {
			  CompleteGrammar G1 = new CompleteGrammar("Grammar"+i, i+"", CC);
			  for (int j = 0; j < 10; j++) {
				  CompleteElementType CX = new CompleteElementType("Struc"+j, G1);
				G1.getSons().add(CX);
			}
			  for (int j = 0; j < 10; j++) {
				  CompleteElementType CX = new CompleteTextElementType("Text"+j, G1);
				G1.getSons().add(CX);
			}
			  for (int j = 0; j < 10; j++) {
				  CompleteElementType CX = new CompleteLinkElementType("Link"+j, G1);
				G1.getSons().add(CX);
			}
			  for (int j = 0; j < 10; j++) {
				  CompleteElementType CX = new CompleteResourceElementType("Struct"+j, G1);
				G1.getSons().add(CX);
			}
			  CC.getMetamodelGrammar().add(G1);
		}
		 
		  
		  
		  processCompleteCollection(new CompleteCollectionLog(), CC, true, System.getProperty("user.home"));
		  
	    }



}
