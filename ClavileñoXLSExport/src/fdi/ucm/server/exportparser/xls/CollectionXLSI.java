/**
 * 
 */
package fdi.ucm.server.exportparser.xls;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
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
import fdi.ucm.server.modelComplete.collection.document.CompleteLinkElement;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElementFile;
import fdi.ucm.server.modelComplete.collection.document.CompleteResourceElementURL;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteLinkElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteResourceElementType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;

/**
 * @author Joaquin Gayoso-Cabada
 *Clase qie produce el XLSI
 */
public class CollectionXLSI {


	public static String processCompleteCollection(CompleteLogAndUpdates cL,
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
        
        HashMap<Long, Integer> clave=new HashMap<Long, Integer>();	
        
//        Sheet hoja;
        
        for (CompleteGrammar row : salvar.getMetamodelGrammar()) {
			processGrammar(libro,row,clave,cL,salvar.getEstructuras(),soloEstructura);
		}
        
        
        libro.write(archivo);

        archivo.close();

		return rutaArchivo;

        

        
        
	}
	
	  private static void processGrammar(Workbook libro, CompleteGrammar grammar,
			HashMap<Long, Integer> clave, CompleteLogAndUpdates cL, List<CompleteDocuments> list, boolean soloEstructura) {
		  
		   Sheet hoja;
		   
		   
		if (!grammar.getNombre().isEmpty())
	        	 {
					int indice = 0;
					String Nombreactual=grammar.getNombre();
					while (libro.getSheet(Nombreactual)!=null)
						{
						Nombreactual=grammar.getNombre()+indice;
						indice++;
						}
					hoja = libro.createSheet(Nombreactual);
	        	 }
	        else hoja = libro.createSheet();
	  
	        
	        List<CompleteElementType> ListaElementos=generaLista(grammar);
	        

	        if (ListaElementos.size()>255)
	        	{
	        	cL.getLogLines().add("Tamaño de estructura demasiado grande para exportar a xls para gramatica: " + grammar.getNombre() +" solo 255 estructuras seran grabadas, divide en gramaticas mas simples");
	        	ListaElementos=ListaElementos.subList(0, 254);
	        	}
	        
	        List<CompleteDocuments> ListaDocumentos=generaDocs(list,grammar);
	      
	        if (ListaDocumentos.size()+1>65536)
	    	{
	    	cL.getLogLines().add("Tamaño de los objetos demasiado grande para exportar a xls");
	    	ListaDocumentos=ListaDocumentos.subList(0, 65536);
	    	}

	        	
	        int row=0;
	        int Column=0;
	        int columnsMax=ListaElementos.size();
	       	
	        
	        for (int i = 0; i < 1; i++) {
	        	Row fila = hoja.createRow(row);
	        	row++;
	        	
	        	for (int j = 0; j < columnsMax+2; j++) {
	        		
	        		String Value = "";
	            	if (j==0)
	            		Value="Id Link Renference";
	            	else 
	            		if (j==1)
	            			Value="Description";
	            		else
	            		{
	            		CompleteElementType TmpEle = ListaElementos.get(j-2);
	            		Value=pathFather(TmpEle);
	            		}
	
	            	
	            	if (Value.length()>=32767)
	            	{
	            		cL.getLogLines().add("Tamaño de Texto en Valor del path del Tipo " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		Value.substring(0, 32766);
	            	}
	            		Cell celda = fila.createCell(j);
	            		
	            		
	            	if (j>1)
	            		{
	            		if (i==0)
	            			clave.put(ListaElementos.get(j-2).getClavilenoid(), Column);
	            		Column++;
	            		}
	            	else
	            	if (j==0)
	            	{
	            		hoja.setColumnWidth(j, 4000);
	            	}
	            	else
	            	if (j==1)
		            {
		            	hoja.setColumnWidth(j, 12750);
		            }
	            	
	            	celda.setCellValue(Value);
	            
	           }
			}	
	        
	        
	        if (!soloEstructura)
	        {
	        /*Hacemos un ciclo para inicializar los valores de filas de celdas*/
	        for(int f=0;f<ListaDocumentos.size();f++){
	            /*La clase Row nos permitirá crear las filas*/
	            Row fila = hoja.createRow(row);
	            row++;

	            CompleteDocuments Doc=ListaDocumentos.get(f);
	            HashMap<Integer, ArrayList<CompleteElement>> ListaClave=new HashMap<Integer, ArrayList<CompleteElement>>();
	            
	            for (CompleteElement elem : Doc.getDescription()) {
					Integer val=clave.get(elem.getHastype().getClavilenoid());
					if (val!=null)
						{
						ArrayList<CompleteElement> Lis=ListaClave.get(val);
						if (Lis==null)
							{
							Lis=new ArrayList<CompleteElement>();
							}
						Lis.add(elem);
						ListaClave.put(val, Lis);
						}
				}
	            
	            
	            
	            /*Cada fila tendrá celdas de datos*/
	            for(int c=0;c<columnsMax+2;c++){
	            	
	            	String Value = "";
	            	if (c==0)
	            		Value=Long.toString(Doc.getClavilenoid());
	            	else if (c==1)
	            		Value=Doc.getDescriptionText();
	            	else
	            		{
	            		ArrayList<CompleteElement> temp = ListaClave.get(c-2);
	            		if (temp!=null)
	            		{
	            		for (CompleteElement completeElement : temp) {
	            			if (!Value.isEmpty())
	            				Value=Value+" "; 
	            			Value=Value+getValueFromElement(completeElement);
								
						}
	            		}
	            		}
	
	            	 
	            	if (Value.length()>=32767)
	            	{
	            		Value="";
	            		cL.getLogLines().add("Tamaño de Texto en Valor en elemento " + Value + " excesivo, no debe superar los 32767 caracteres, columna recortada");
	            		Value.substring(0, 32766);
	            	}
	                /*Creamos la celda a partir de la fila actual*/
	                Cell celda = fila.createCell(c);               	
	                		 celda.setCellValue(Value);
	                    /*Si no es la primera fila establecemos un valor*/
	                	//32.767

	                
	            	}

	            		
	            		
	            }
	        
	        }
	        
	       
		
	}

	private static ArrayList<CompleteDocuments> generaDocs(
			List<CompleteDocuments> list, CompleteGrammar grammar) {
		ArrayList<CompleteDocuments> ListaDoc=new ArrayList<CompleteDocuments>();
		for (CompleteDocuments completeDocuments : list) {
			if (StaticFuctionsXLS.isInGrammar(completeDocuments,grammar))
				ListaDoc.add(completeDocuments);
		}
		return ListaDoc;
	}

//	private static ArrayList<CompleteElementType> generaLista(
//			List<CompleteGrammar> metamodelGrammar) {
//		  ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
//		  for (CompleteGrammar completegramar : metamodelGrammar) {
//			ListaElementos.addAll(generaLista(completegramar));
//		}
//		return ListaElementos;
//	}

	private static ArrayList<CompleteElementType> generaLista(
			CompleteGrammar completegramar) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 for (CompleteElementType completeelem : completegramar.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType
			 				&&(!StaticFuctionsXLS.isIgnored((CompleteElementType)completeelem)))
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
				ListaElementos.addAll(generaLista(completeelem));
			}
		 return ListaElementos;
	}

	private static Collection<? extends CompleteElementType> generaLista(
			CompleteElementType completeelementPadre) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 for (CompleteElementType completeelem : completeelementPadre.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType
			 				&&(!StaticFuctionsXLS.isIgnored((CompleteElementType)completeelem)))
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
				ListaElementos.addAll(generaLista(completeelem));
			}
		 return ListaElementos;
	}

	private static String getValueFromElement(CompleteElement completeElement) {
		try {
			if (completeElement instanceof CompleteTextElement)
    			return (((CompleteTextElement)completeElement).getValue());
			else if (completeElement instanceof CompleteLinkElement)
				return Long.toString((((CompleteLinkElement)completeElement).getValue().getClavilenoid()));
			else if (completeElement instanceof CompleteResourceElementURL)
				return (((CompleteResourceElementURL)completeElement).getValue());
			else if (completeElement instanceof CompleteResourceElementFile)
				return (((CompleteResourceElementFile)completeElement).getValue().getPath());
		} catch (Exception e) {
			return "";
		}
		return "";
	}
	
	public static void main(String[] args) throws Exception{
	
		
		String message="Exception .clavy-> Params Null ";
		try {

			
			
			String fileName = "test.clavy";
			
			if (args.length!=0)
				fileName=args[0];
			
			boolean soloestruct=false;
			if (args.length>1)
				try {
					soloestruct=Boolean.parseBoolean(args[1]);
				} catch (Exception e) {
					
				}
			
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
			 
			 
		
		 
		  
		  
			 processCompleteCollection(new CompleteLogAndUpdates(), object, soloestruct, System.getProperty("user.home"));
			 
	    }catch (Exception e) {
			e.printStackTrace();
			System.err.println(message);
			throw new RuntimeException(message);
		}
		  
		  
		 
		  
	    }


	/**
	 *  Retorna el Texto que representa al path.
	 *  @return Texto cadena para el elemento
	 */
	public static String pathFather(CompleteElementType entrada)
	{
		String DataShow= ((CompleteElementType) entrada).getName();

		
		if (entrada.getFather()!=null)
			return pathFather(entrada.getFather())+"/"+DataShow;
		else return DataShow;
	}
	
}
