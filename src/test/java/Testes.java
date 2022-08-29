import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.basic.*;
import com.automationanywhere.botcommand.samples.commands.basic.DataFormat;
import com.automationanywhere.botcommand.samples.commands.conditionals.SheetExists;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
//import org.apache.poi.hssf.model.InternalWorkbook;
//import org.apache.poi.xssf.model.*;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.ExternalLinksTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;

public class Testes {

    public void ttt() {
        Open o = new Open();
        o.execute("");
        System.out.println("asd".matches("a"));

    }

    @Test
    public void test(){
        String inputFile = "C:\\Users\\melque\\Documents\\WEEKLY FBDNPC - SEM 3.xlsx";
        String outputFile = "C:\\Users\\melque\\Documents\\testiiiiiiiiii.xls";
        xls2xlsx a = new xls2xlsx();

        a.action(inputFile,outputFile);
    }

    public void tt(){
        String inputFile = "C:\\Users\\melque\\Documents\\WEEKLY FBDNPC - SEM 3.xlsx";
        DataFormat a = new DataFormat();
        a.action(
                inputFile,
                "name",
                "QUADRO",
                0.0,
                "B5:AL1000",
                false,
                "custom",
                "_-R$ * #,###.##;_-@_-",
                false,
                "RIGHT",
                false,
                "BOTTOM",
                true,
                false
        );
        a.action(
                inputFile,
                "name",
                "QUADRO",
                0.0,
                "B5:AL109",
                false,
                "custom",
                "_-R$ * #,###.##;_-@_-",
                false,
                "RIGHT",
                false,
                "BOTTOM",
                true,
                true
        );
//        a.action(
//                inputFile,
//                "name",
//                "Desligados",
//                0.0,
//                "C2:C100",
//                true,
//                "14",
//                "",
//                false,
//                "RIGHT",
//                false,
//                "BOTTOM"
//        );
//
        //================================================================= CREATE WORKBOOK OBJECT
//        WorkbookHelper wbH = new WorkbookHelper(inputFile);
//
//        //================================================================= GET SHEET
//        Sheet mySheet = wbH.wb.getSheet("Desligados");
//        Cell cl2 = mySheet.getRow(0).getCell(1);
//        CellStyle CellStyle2 = wbH.wb.createCellStyle();
//        CellStyle2.cloneStyleFrom(cl2.getCellStyle());
//        CellStyle2.setDataFormat((short) 0);
//        cl2.setCellStyle(CellStyle2);


//        CellRangeAddress aa = CellRangeAddress.valueOf("A2:A17");
//        for(CellAddress d: aa){
//            Cell cl = wbH.getCell(mySheet,d);
//            if (cl != null) {
//                CellStyle CellStyle = cl.getCellStyle();
//                //System.out.println("FORMAT:" + CellStyle.getAlignment() + ":" + CellStyle.getVerticalAlignment());
//                CellStyle.setDataFormat((short) 14);
//                //cl.setCellStyle(CellStyle);
//
//                //System.out.println("STR:" + d.toString() + " C:" + d.getColumn() + " R:" + d.getRow());
//                System.out.println(cl);
//            }
//        }
        //wbH.save(inputFile);

    }

   public void t(){
        LockUnlockCells a = new LockUnlockCells();
        SheetExists b = new SheetExists();
        GetSheets c = new GetSheets();

        String inputFile = "C:\\Users\\melque\\Documents\\WEEKLY FBDNPC - SEM 3.xlsx";


        //System.out.println(b.validate(inputFile,"not in","AGN"));
        //System.out.println(c.action(inputFile).get().size());
    a.action(inputFile,"name","QUADRO",null,"B5:AL1000","unlock");
    //a.action(inputFile,"name","QUADRO",null,"B5:AL109","lock");
    }


    public void s(){
        GetReferences gr= new GetReferences();
        //String inputFile = "C:\\Users\\melque\\Documents\\teste3.xls";
        String inputFile = "C:\\Users\\melque\\Documents\\RF21 - PREMISSAS - AGNIVJU.xlsx";
        //WorkbookHelper wbh = new WorkbookHelper(inputFile);
        //LinkedList<String> a = getWorkbookReferences((XSSFWorkbook) wbh.wb);
        ListValue<String> lista = gr.action(inputFile);
        System.out.println(lista.get(0));
    }

    private LinkedList<String> getWorkbookReferences(HSSFWorkbook wb) {
        LinkedList<String> references = new LinkedList<>();

        try {
            // 1. Get InternalWorkbook
            Field internalWorkbookField = HSSFWorkbook.class.getDeclaredField("workbook");
            internalWorkbookField.setAccessible(true);
            InternalWorkbook internalWorkbook = wb.getInternalWorkbook();

            // 2. Get LinkTable (hidden class)
            Method getLinkTableMethod;
            getLinkTableMethod = InternalWorkbook.class.getDeclaredMethod("getOrCreateLinkTable", null);

            getLinkTableMethod.setAccessible(true);
            Object linkTable = getLinkTableMethod.invoke(internalWorkbook, null);

            // 3. Get external books method
            Method externalBooksMethod = linkTable.getClass().getDeclaredMethod("getExternalBookAndSheetName", int.class);
            externalBooksMethod.setAccessible(true);

            // 4. Loop over all possible workbooks
            int i = 0;
            String[] names;
            try {
                while( true) {
                    names = (String[]) externalBooksMethod.invoke(linkTable, i++) ;                     if (names != null ) {
                        references.add(names[0]);
                    }
                }
            }
            catch  ( java.lang.reflect.InvocationTargetException e) {
                if ( !(e.getCause() instanceof java.lang.IndexOutOfBoundsException) ) {
                    throw e;
                }
            }
        } catch (NoSuchFieldException | NoSuchMethodException | SecurityException | InvocationTargetException | IllegalAccessException e) {
            e.printStackTrace();
        }

        return references;
    }

    private LinkedList<String> getWorkbookReferences(XSSFWorkbook wb) {
        LinkedList<String> references = new LinkedList<>();

        try {
            // 1. Get InternalWorkbook
            Field internalWorkbookField = XSSFWorkbook.class.getDeclaredField("workbook");
            internalWorkbookField.setAccessible(true);

            List<ExternalLinksTable> linkTable = wb.getExternalLinksTable();
            System.out.println(linkTable.size());

            List<String> wbks = new ArrayList<>();

            for(ExternalLinksTable links : linkTable){
                wbks.add(links.getLinkedFileName().replaceAll("%20"," ").replace("file:///",""));
            }

            wbks = wbks.stream().sorted().collect(Collectors.toList());

            for(String a: wbks){
                System.out.println(a);
            }
            System.out.println();

        } catch (NoSuchFieldException | SecurityException e) {
            e.printStackTrace();
        }

        return references;
    }


}
