/**
 * 
 */
package fdi.ucm.server.exportparser.xls.properties;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Properties;
import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.CompleteLogAndUpdates;
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
        String rutaArchivo = pathTemporalFiles+"/"+System.nanoTime()+".properties";
        /*Se crea el objeto de tipo File con la ruta del archivo*/
        File archivoProperties = new File(rutaArchivo);
        /*Si el archivo existe se elimina*/
        if(archivoProperties.exists()) archivoProperties.delete();
        /*Se crea el archivo*/
        archivoProperties.createNewFile();
        

      Properties Resultado=new Properties();
        
      HashMap<CompleteGrammar, String> NombreReal=new HashMap<CompleteGrammar, String>();
      
      HashMap<String, CompleteGrammar> NombreRealSime=new HashMap<String, CompleteGrammar>();
      
      HashMap<String, Integer> NombreRealSimeCount=new HashMap<String, Integer>();
      
      for (CompleteGrammar row : salvar.getMetamodelGrammar()) {
    	  
    	  String NombreActual = row.getNombre();
    	  
    	  Integer I = NombreRealSimeCount.get(NombreActual);
    	  
    	  if (I!=null)
    	  {
    		  CompleteGrammar CC1= NombreRealSime.get(NombreActual);
    		  NombreReal.put(CC1, NombreActual+"#1");
    		  
    		  NombreRealSime.put(NombreActual+"#1", CC1);
    		 
    		  I = new Integer(I.intValue()+1);
    		  
    		  NombreReal.put(row, NombreActual+"#"+I.intValue());
    		  NombreRealSime.put(NombreActual+"#"+I.intValue(), row);
    		  
    		 
    		  
    	  }else
    		  {
    		  I = new Integer(1);
    		  NombreRealSime.put(NombreActual, row);
    		  }
    		  
    		 
    	  
    	  NombreRealSimeCount.put(NombreActual, I);
    	  

      }
      
//        Sheet hoja;
        
        for (CompleteGrammar row : salvar.getMetamodelGrammar()) {
        	
        	String StringGramG = NombreReal.get(row);
        	
        	if (StringGramG==null)
        		StringGramG=row.getNombre();
        	
			processGrammar(row,Resultado,cL,StringGramG);
		}
        
        Resultado.store(new FileOutputStream(archivoProperties), null);

		return rutaArchivo;

        

        
        
	}

	private static void processGrammar(CompleteGrammar grammar,
			Properties c, CompleteLogAndUpdates cL, String stringGramG) {
		  
	
		HashMap<CompleteElementType,String> TablaNombres=new HashMap<CompleteElementType, String>();
		
	        List<CompleteElementType> ListaElementos=generaLista(grammar,TablaNombres);
	        
	        


	        	
	        	for (int j = 0; j < ListaElementos.size(); j++) {
	        		
	        		Long Value1 = -1L ;
	        		String Value2 = "~uname";
	        		
	
	            		CompleteElementType TmpEle = ListaElementos.get(j);
	            		Value2=stringGramG+"/"+pathFather(TmpEle,TablaNombres);
	            		Value1=TmpEle.getClavilenoid();
	
	            		c.setProperty(Value1.toString(),Value2);
	            		
	            	
	            	
	           }
			
	        
	       
		
	}

	


	private static ArrayList<CompleteElementType> generaLista(
			CompleteGrammar completegramar, HashMap<CompleteElementType,String> TablaNombres) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 HashMap<Long,Integer> TablaNombresConteo =new HashMap<Long, Integer>();
		 for (CompleteElementType completeelem : completegramar.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType
			 				&&(!StaticFuctionsXLS.isIgnored((CompleteElementType)completeelem)))
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
			 	
			 	if (completeelem.getClassOfIterator()!=null&&completeelem.getClassOfIterator().isMultivalued())
			 	{
			 		Integer AA=TablaNombresConteo.get(completeelem.getClassOfIterator().getClavilenoid());
			 		if (AA==null)
			 			AA=new Integer(0);
			 		
			 		AA=new Integer(AA.intValue()+1);
			 		
			 		TablaNombres.put(completeelem, completeelem.getClassOfIterator().getName()+"#"+AA.intValue());
			 		
			 		TablaNombresConteo.put(completeelem.getClassOfIterator().getClavilenoid(), AA);
			 		
			 	}
				ListaElementos.addAll(generaLista(completeelem,TablaNombres));
			}
		 return ListaElementos;
	}

	private static Collection<? extends CompleteElementType> generaLista(
			CompleteElementType completeelementPadre, HashMap<CompleteElementType,String> TablaNombres) {
		 ArrayList<CompleteElementType> ListaElementos = new ArrayList<CompleteElementType>();
		 HashMap<Long,Integer> TablaNombresConteo =new HashMap<Long, Integer>();
		 for (CompleteElementType completeelem : completeelementPadre.getSons()) {
			 	if (completeelem instanceof CompleteElementType)
			 		{
			 		if (completeelem instanceof CompleteTextElementType||completeelem instanceof CompleteLinkElementType||completeelem instanceof CompleteResourceElementType
			 				&&(!StaticFuctionsXLS.isIgnored((CompleteElementType)completeelem)))
			 			ListaElementos.add((CompleteElementType)completeelem);
			 		}
			 	
			 	
			 	if (completeelem.getClassOfIterator()!=null&&completeelem.getClassOfIterator().isMultivalued())
			 	{
			 		Integer AA=TablaNombresConteo.get(completeelem.getClassOfIterator().getClavilenoid());
			 		if (AA==null)
			 			AA=new Integer(0);
			 		
			 		AA=new Integer(AA.intValue()+1);
			 		
			 		TablaNombres.put(completeelem, completeelem.getClassOfIterator().getName()+"#"+AA.intValue());
			 		
			 		TablaNombresConteo.put(completeelem.getClassOfIterator().getClavilenoid(), AA);
			 		
			 	}
			 	
				ListaElementos.addAll(generaLista(completeelem,TablaNombres));
			}
		 return ListaElementos;
	}

	
	public static void main(String[] args) throws Exception{
	
		System.out.println("properties_V1");
		String message="Exception .clavy-> Params Null ";
		try {

			
			
			String fileName = "test.clavy";
			
			if (args.length!=0)
				fileName=args[0];
			
			
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
			 
			 
		
		 
		  
		  
			 processCompleteCollection(new CompleteLogAndUpdates(), object, true, System.getProperty("user.home"));
			 
	    }catch (Exception e) {
			e.printStackTrace();
			System.err.println(message);
			throw new RuntimeException(message);
		}
		  
		  
		 
		  
	    }


	/**
	 *  Retorna el Texto que representa al path.
	 * @param tablaNombres 
	 *  @return Texto cadena para el elemento
	 */
	public static String pathFather(CompleteElementType entrada, HashMap<CompleteElementType, String> tablaNombres)
	{
		String DataShow=  entrada.getName();
		
		if (tablaNombres.get(entrada)!=null)
			DataShow = tablaNombres.get(entrada);

		
		if (entrada.getFather()!=null)
			return pathFather(entrada.getFather(),tablaNombres)+"/"+DataShow;
		else return DataShow;
	}
	
}
