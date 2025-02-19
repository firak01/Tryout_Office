package use.office.libre.search;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

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
                                searchContent(file);
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
    //  is_odt
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

    //
    // search
    //
    private void searchContent(File odt) {
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
