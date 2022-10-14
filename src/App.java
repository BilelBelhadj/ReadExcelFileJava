/*
    Titre           : Real excel file
    Description     : le programme permet de lire les donnees d'un fichier excel existant et l'affiche dans le moniteur serie
    Auteur          : Bilel Belhadj
    Version         : 0.0.1

 
 
 */


//importer les fichiers necessaire 
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class App {
    public static void main(String[] args) throws Exception {
        
    //declaration de variable de type file et specefier le chemin
    File myFile = new File("C:/Users/PC/Desktop/CCNB/2eme/PROG1284/JavaExcelFile/staff.xlsx");
            FileInputStream fis = new FileInputStream(myFile);

            // creer un istence XSSFWorkbook pour le pouvoir manipuler das le program java
            XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
           
            // creer un instence XSSFsheet 
            XSSFSheet mySheet = myWorkBook.getSheetAt(0);
           
            // variable de maipulation des lignes de fichier excel
            Iterator<Row> rowIterator = mySheet.iterator();
           
            // verifier si le ligne pas vide
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // lire les colomne
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {

                    //lire le cotenue de la colomne das le variale
                    Cell cell = cellIterator.next();

                    //affichage des cellules
                    switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "\t");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t");
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue() + "\t");
                        break;
                    default :
                 
                    }
                }
                System.out.println("");
            }   

    }
}
