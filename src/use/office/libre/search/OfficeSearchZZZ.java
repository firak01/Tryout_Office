package use.office.libre.search;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import org.jdom2.Attribute;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.JDOMException;
import org.jdom2.Namespace;
import org.jdom2.filter.Filters;
import org.jdom2.input.SAXBuilder;

/**
 *  OfficeSearch
 *  siehe: https://stackoverflow.com/questions/54293459/find-a-string-in-a-list-of-odt-files-and-print-the-matching-lines
 */
public class OfficeSearchZZZ{
    private final Set<String> searchSet = new HashSet<>();
    private static OfficeSearchZZZ INSTANCE = new OfficeSearchZZZ();
    
    ArrayList<String>listasFile=null;
    
    //
    // main
    //
    public static void main(String[] args) {
        INSTANCE.execute(args);
    }

    //
    //  execute
    //
    private void execute(String[] args) {
        if (args.length <= 1) {        	 
                 System.out.println("Usage: OfficeSearchZZZ <directory> <search_term> [...]");
             }else {
        	
            for (int i=1; i<args.length; i++) {
                searchSet.add(args[i].toLowerCase());
            }
//            try {
            	//Java 1.8 Syntax
            	/*
                Files.list(Paths.get(args[0])).sorted().
                    map(Path::toFile).
                    filter(this::isOdt).
                    forEach(this::search);
                 */
            	
            	//Umwandlung in Java 1.7. Syntax (ohne Streams-API und Files.list-Methode
            	File directory = new File(args[0]);
                if (directory.isDirectory()) {
                    File[] files = directory.listFiles();
                    if (files != null) {
                        // Sortiere die Dateien alphabetisch
                        Arrays.sort(files);

                        for (File file : files) {
                            if (isOdt(file)) {
                                //searchContentBodyText(file);
                            	//searchContentStyle(file);
                            	searchStylesStyle(file, false, false);
                            }
                        }
                    } else {
                        System.out.println("Keine Dateien gefunden.");
                    }
                } else {
                    System.out.println("Das angegebene Argument ist kein Verzeichnis.");
                }
            }
            
//            catch (IOException e) {
//                e.printStackTrace();
//            }
        }
       

    //
    //  isOdt
    //
    private boolean isOdt(File file) {
        if (file.isFile()) {
            final String name = file.getName();
            final int dotidx = name.lastIndexOf('.');
            if ((0 <= dotidx) && (dotidx < name.length() - 1)) {
                return name.substring(dotidx + 1).equalsIgnoreCase("odt");
            }
        }
        return false;
    }
    
    public ArrayList<String> getInternalFilesSearchedFor(){
    	if(listasFile==null) {
    		listasFile=new ArrayList<String>();
    		listasFile.add("content.xml");
    	}    	
    	return this.listasFile;
    }
    
    public void setFilesSearchedFor(ArrayList<String> listasFile) {
    	this.listasFile = listasFile;
    }

    private void searchStylesStyle(File odt) {
    	searchStylesStyle_(odt, true, true);
    }
    
    private void searchStylesStyle(File odt, boolean bContainsPhrase) {
    	searchStylesStyle_(odt, bContainsPhrase, true);
    }
    
    private void searchStylesStyle(File odt, boolean bContainsPhrase, boolean bExact ) {
    	searchStylesStyle_(odt, bContainsPhrase, bExact);
    }
    
    private void searchStylesStyle_(File odt, boolean bContainsPhrase, boolean bExact ) {
    	ArrayList<String>listasFile=this.getInternalFilesSearchedFor();
    			
        try (ZipFile zip = new ZipFile(odt)) {
        	
        			
            final ZipEntry content = zip.getEntry("content.xml");
            if (content != null) {
            	boolean bFoundAtFile = false;
            	boolean bFoundAtElement = false;
            	
                final SAXBuilder builder = new SAXBuilder();
                final Document doc = builder.build(zip.getInputStream(content));
                final Element root = doc.getRootElement();
                final Namespace office_ns = root.getNamespace("office");
                final Namespace style_ns = root.getNamespace("style");
                final Element automaticStyles = root.getChild("automatic-styles", office_ns);
                if (automaticStyles != null) {                    
                    for (Element e : automaticStyles.getDescendants(Filters.element(style_ns))) {
                    		//final String sName = e.getName();  
                    		//System.out.print("\n" + odt.toString() + " : Attribut: " + sName);
                    				
                    		//String sNameSpace = office_ns.getPrefix();
                    		String sNameSpace = style_ns.getPrefix();
                    		List<Attribute> listAttribute = e.getAttributes();
                    		for(Attribute objAttribute: listAttribute) {                              
                    			String sAttributeName = objAttribute.getName();
                    			if(sAttributeName.equals("name")) {   
                    				String sAttributeValue = objAttribute.getValue();
	                    			for (String p : searchSet) {	   
	                    				if(bContainsPhrase) {
	                    					if(bExact) {
				                                if (sAttributeValue.toLowerCase().contains(p)) {
				                                        bFoundAtElement = true;
				                                        bFoundAtFile = true;		                                       		                                       
				                                }
	                    					}else {
	                    						if (sAttributeValue.contains(p)) {
			                                        bFoundAtElement = true;
			                                        bFoundAtFile = true;		                                       		                                       
	                    						}
			                                }
	                    				}else {
	                    					//!bContainsPhrase
	                    					if(bExact) {
				                                if (sAttributeValue.toLowerCase().equals(p)) {
				                                        bFoundAtElement = true;
				                                        bFoundAtFile = true;		                                       		                                       
				                                }
	                    					}else {
	                    						if (sAttributeValue.equals(p)) {
			                                        bFoundAtElement = true;
			                                        bFoundAtFile = true;		                                       		                                       
	                    						}
			                                }
	                    				}
		                                
		                                if(bFoundAtElement) {
		                                	System.out.print("\n!!! " + odt.toString() + " : !!! Attribut: " + sAttributeName + "\t| Wert: " + sAttributeValue +" | found: " + p);
		                                	bFoundAtElement=false;
		                                }else {
		                                	System.out.print("\n" + odt.toString() + " : Attribut: " + sAttributeName + "\t| Wert: " + sAttributeValue + " | but searched for: " + p);
		                                }
		                            }//end for
                    			}//if(sAttributeName.equals("name")) {                    			                    		
                    		}//end for                    		                    		                     
                    }
                }else{
                	System.out.println("\n" + odt.toString() + " enthält keinen Knoten 'automatic-styles'.");                	
                }//end if "Knoten in Struktur" gefunden.
                
                if(bFoundAtFile) {
                	System.out.println("\n!!!" + odt.toString() + " enthält gesuchte(s) Attribut(e).");
                }
            }else {
            	System.out.println("\n" + odt.toString() + " enthält keine Datei 'content.xml'.");
            }//end if content!=null, d.h. "Datei" gefunden.
        }
        catch (IOException | JDOMException e) {
            e.printStackTrace();
        }
    }

    
    private void searchContentStyle(File odt) {
        try (ZipFile zip = new ZipFile(odt)) {
            final ZipEntry content = zip.getEntry("content.xml");
            if (content != null) {
            	boolean bFoundAtFile = false;
            	boolean bFoundAtElement = false;
            	
                final SAXBuilder builder = new SAXBuilder();
                final Document doc = builder.build(zip.getInputStream(content));
                final Element root = doc.getRootElement();
                final Namespace office_ns = root.getNamespace("office");
                final Namespace style_ns = root.getNamespace("style");
                final Element automaticStyles = root.getChild("automatic-styles", office_ns);
                if (automaticStyles != null) {                    
                    for (Element e : automaticStyles.getDescendants(Filters.element(style_ns))) {
                    		//final String sName = e.getName();  
                    		//System.out.print("\n" + odt.toString() + " : Attribut: " + sName);
                    				
                    		//String sNameSpace = office_ns.getPrefix();
                    		String sNameSpace = style_ns.getPrefix();
                    		List<Attribute> listAttribute = e.getAttributes();
                    		for(Attribute objAttribute: listAttribute) {                              
                    			String sAttributeName = objAttribute.getName();
                    			if(sAttributeName.equals("name")) {                    				
	                    			for (String p : searchSet) {	                    					                    				
		                                if (sAttributeName.toLowerCase().contains(p)) {
		                                        bFoundAtElement = true;
		                                        bFoundAtFile = true;		                                       		                                       
		                                }
		                                
		                                String sValue = objAttribute.getValue();
		                                if(bFoundAtElement) {
		                                	System.out.print("\n" + odt.toString() + " : !!! Attribut: " + sAttributeName + "\t| Wert: " + sValue);
		                                	bFoundAtElement=false;
		                                }else {
		                                	System.out.print("\n" + odt.toString() + " : Attribut: " + sAttributeName + "\t| Wert: " + sValue);
		                                }
		                            }
                    			}
                    			
                    			
                    		}
                    		
                    		
                     
                    }
                }else{
                	System.out.println("\n" + odt.toString() + " enthält keinen Knoten 'automatic-styles'.");                	
                }//end if "Knoten in Struktur" gefunden.
                
                if(bFoundAtFile) {
                	System.out.println("\n!!!" + odt.toString() + " enthält gesuchte(s) Attribut(e).");
                }
            }else {
            	System.out.println("\n" + odt.toString() + " enthält keine Datei 'content.xml'.");
            }//end if content!=null, d.h. "Datei" gefunden.
        }
        catch (IOException | JDOMException e) {
            e.printStackTrace();
        }
    }

    
    //
    // search
    //
    private void searchContentBodyText(File odt) {
        try (ZipFile zip = new ZipFile(odt)) {
            final ZipEntry content = zip.getEntry("content.xml");
            if (content != null) {
                final SAXBuilder builder = new SAXBuilder();
                final Document doc = builder.build(zip.getInputStream(content));
                final Element root = doc.getRootElement();
                final Namespace office_ns = root.getNamespace("office");
                final Namespace text_ns = root.getNamespace("text");
                final Element body = root.getChild("body", office_ns);
                if (body != null) {
                    boolean found = false;
                    for (Element e : body.getDescendants(Filters.element(text_ns))) {
                        if ("p".equals(e.getName()) ||
                            "h".equals(e.getName())) {
                            final String s = e.getValue().toLowerCase();
                            for (String p : searchSet) {
                                if (s.contains(p)) {
                                    if (!found) {
                                        found = true;
                                        System.out.println("\n" + odt.toString());
                                    }
                                    System.out.println(e.getValue());
                                }
                            }
                        }
                    }
                }
            }
        }
        catch (IOException | JDOMException e) {
            e.printStackTrace();
        }
    }
    
    
   

}
