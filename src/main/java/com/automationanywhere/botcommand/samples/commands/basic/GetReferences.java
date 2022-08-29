package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.GreaterThanEqualTo;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.ExternalLinksTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;

import static com.automationanywhere.commandsdk.model.AttributeType.*;

//BotCommand makes a class eligible for being considered as an action.
//@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "GetReferences",
        label = "GetReferences",
        node_label = "Get references from {{file}}",
        description = "",
        icon = "pkg.svg",
        return_type = DataType.LIST,
        return_required = true
)

public class GetReferences {
    @Execute
    public ListValue action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Excel file",description = "example: C:\\folder\\file.xlsx")
            //@FileExtension("xlsx")
            @NotEmpty
            String file
) {
        WorkbookHelper wbh = new WorkbookHelper(file);
        List<String> lista = new LinkedList<>();

        if(file.toUpperCase().endsWith(".XLSX")) {
            lista = getWorkbookReferences((XSSFWorkbook) wbh.wb);
        }else{
            lista = getWorkbookReferences((HSSFWorkbook) wbh.wb);
        }
        List<Value> OutPut = new ArrayList<Value>();
        for(String a: lista){
            OutPut.add(new StringValue(a));
        }
        ListValue rtrn =  new ListValue();
        rtrn.set(OutPut);
        return rtrn;
    }

    private List<String> getWorkbookReferences(HSSFWorkbook wb) {
        List<String> references = new ArrayList<>();

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
            throw new BotCommandException("Error Getting references: " + e.getMessage());
        }

        references = references.stream().sorted().collect(Collectors.toList());
        return references;
    }

    private List<String> getWorkbookReferences(XSSFWorkbook wb) {
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

//            for(String a: wbks){
//                System.out.println(a);
//            }
//            System.out.println();
            return wbks;

        } catch (NoSuchFieldException | SecurityException e) {
            throw new BotCommandException("Error Getting references: " + e.getMessage());
        }
    }


}
