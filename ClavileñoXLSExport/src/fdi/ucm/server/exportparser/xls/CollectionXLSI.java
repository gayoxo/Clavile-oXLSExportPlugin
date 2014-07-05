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
import java.util.Random;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.CompleteCollectionLog;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteFile;
import fdi.ucm.server.modelComplete.collection.document.CompleteLinkElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElementFile;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElementURL;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
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
        
        boolean Errores=false;
        
        if (ListaElementos.size()>255)
        	{
        	cL.getLogLines().add("Tamaño de estructura demasiado grande para exportar a xls");
        	Errores=true;
        	}
        	
        if (salvar.getEstructuras().size()+1>65536)
    	{
    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls");
    	Errores=true;
    	}
       
        if (!Errores)
        {
        /*Hacemos un ciclo para inicializar los valores de 10 filas de celdas*/
        for(int f=0;f<(salvar.getEstructuras().size()+2);f++){
            /*La clase Row nos permitirá crear las filas*/
            Row fila = hoja.createRow(f);

            /*Cada fila tendrá 5 celdas de datos*/
            for(int c=0;c<=ListaElementos.size();c++){
            	
            	String Value = "";
            	if (c!=0)
            		Value=ListaElementos.get(c-1).getName();
            	
            	String Value2 = "";
            	if (c!=0)
            		Value2=ListaElementos.get(c-1).getClavilenoid().toString();
	
            	if (Value.length()<32767)
            	{
                /*Creamos la celda a partir de la fila actual*/
                Cell celda = fila.createCell(c);
                
                /*Si la fila es la número 0, estableceremos los encabezados*/
                if(f==0){
                	
                		celda.setCellValue(Value);
                }else if (f==1){
                		celda.setCellValue(Value2);
                }else{
                	
                	if (c==0)
                		celda.setCellValue("Valor Documento "+(f-2));
                	else
                		 celda.setCellValue("Valor celda "+(c-1)+","+(f-2));
                    /*Si no es la primera fila establecemos un valor*/
                	//32.767
                   
                }
                
            	}
            	else 
            		cL.getLogLines().add("Temaño de Texto en Structura " + Value + " excesivo, no debe superar los 32767 caracteres, columna ignorada");
            		
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
        else 
        	{
        	 libro.write(archivo);
        	archivo.close();
        	return "";
        	}
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
		
		int id=0;
		
		
		
		  CompleteCollection CC=new CompleteCollection("Lou Arreglate", "Arreglate ya!");
		  for (int i = 0; i < 5; i++) {
			  CompleteGrammar G1 = new CompleteGrammar(new Long(id),"Grammar"+i, i+"", CC);
			  
			  ArrayList<CompleteDocuments> CD=new ArrayList<CompleteDocuments>();
			  int docsN=(new Random()).nextInt(10);
			for (int j = 0; j < docsN; j++) {
				CompleteDocuments CDDD=new CompleteDocuments(new Long(id), CC, G1, "", "");
				CC.getEstructuras().add(CDDD);
				 id++;
				CD.add(CDDD);
			}
			  
			  id++;
			  for (int j = 0; j < 5; j++) {
				  CompleteElementType CX = new CompleteElementType(new Long(id),"Struc "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteTextElementType CX = new CompleteTextElementType(new Long(id),"Text "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						CompleteTextElement CTE=new CompleteTextElement(new Long(id), CX, "ValorText "+(i*10+j));
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
				
				
				
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteLinkElementType CX = new CompleteLinkElementType(new Long(id),"Link "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						CompleteLinkElement CTE=new CompleteLinkElement(new Long(id), CX, CD.get((new Random()).nextInt(CD.size())));
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
			}
			  for (int j = 0; j < 5; j++) {
				  CompleteResourceElementType CX = new CompleteResourceElementType(new Long(id),"Resour "+(i*10+j), G1);
				  id++;
				G1.getSons().add(CX);
				
				for (CompleteDocuments completeDocuments : CD) {
					boolean docrep=(new Random()).nextBoolean();
					if (docrep)
						{
						
						boolean URL=(new Random()).nextBoolean();
						CompleteResourceElement CTE;
						if (URL)
							CTE=new CompleteResourceElementURL(new Long(id), CX, "ValorText "+(i*10+j));
						else 
							{
							CompleteFile FF = new CompleteFile(new Long(id), "ValorText "+(i*10+j), CC);
							CC.getSectionValues().add(FF);
							id++;
							CTE=new CompleteResourceElementFile(new Long(id), CX, FF);
							}
						id++;
						completeDocuments.getDescription().add(CTE);
						}
				}
				
			}
			  CC.getMetamodelGrammar().add(G1);
		}
		 
		  
		  
		  processCompleteCollection(new CompleteCollectionLog(), CC, true, System.getProperty("user.home"));
		  
	    }



}
