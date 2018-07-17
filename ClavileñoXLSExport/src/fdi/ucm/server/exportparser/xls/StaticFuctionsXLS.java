/**
 * 
 */
package fdi.ucm.server.exportparser.xls;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;

import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteOperationalValueType;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteElementType;

/**
 * Funcion que implementa las funciones estaticas de la exportacion
 * @author Joaquin Gayoso-Cabada
 *
 */
public class StaticFuctionsXLS {
	
	
	public static boolean isInGrammar(CompleteDocuments iterable_element,
			CompleteGrammar completeGrammar) {
		HashSet<Long> ElemT=new HashSet<Long>();
		for (CompleteElement dd : iterable_element.getDescription()) {
			ElemT.add(dd.getHastype().getClavilenoid());
		}
		
		return isInGrammar(ElemT, completeGrammar.getSons());
		
		
	}



	private static boolean isInGrammar(HashSet<Long> elemT,
			List<CompleteElementType> sons) {
		for (CompleteElementType CSlong1 : sons) {
			if (elemT.contains(CSlong1.getClavilenoid())||isInGrammar(elemT, CSlong1.getSons()))
				return true;
			
		}
		return false;
	}
	
	public static boolean isIgnored(CompleteElementType hastype) {
		ArrayList<CompleteOperationalValueType> Shows = hastype.getShows();
		for (CompleteOperationalValueType show : Shows) {
			
			if (show.getView().equals(NameConstantsXLS.META))
			{

					if (show.getName().equals(NameConstantsXLS.TYPE))
						if (show.getDefault().equals(NameConstantsXLS.IGNORED)) 
										return true;

			}
		}
		return false;
	}

}
