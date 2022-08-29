package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.utils.FindInListSchema;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.model.ExternalLinksTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.xml.crypto.Data;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.stream.Collectors;

import static com.automationanywhere.commandsdk.model.AttributeType.FILE;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "GetSheets",
        label = "GetSheets",
        node_label = "Get sheets from {{file}}",
        description = "",
        icon = "pkg.svg",
        return_type = DataType.LIST,
        return_required = true,
        return_sub_type = DataType.STRING
)

public class GetSheets {
    @Execute
    public ListValue action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Excel file",description = "example: C:\\folder\\file.xlsx")
            //@FileExtension("xlsx")
            @NotEmpty
            String file
) {
        Workbook wb = this.wbget(file);
        WorkbookHelper wbh = new WorkbookHelper(wb);
        List<String> lista = wbh.getSheetsName();

        List<Value> OutPut = new ArrayList<Value>();
        for(String a: lista){
            OutPut.add(new StringValue(a));
        }
        ListValue rtrn =  new ListValue();
        rtrn.set(OutPut);
        return rtrn;
    }
    private Workbook wbget(String file){
        try{
            Workbook myWorkBook;
            File fl = new File(file);
            long mb = fl.length()/1024/1024;
            if((mb > 10 && file.toUpperCase().endsWith(".XLSX")) || file.toUpperCase().endsWith(".XLSX")){
                FileInputStream is = new FileInputStream(fl);
                myWorkBook = StreamingReader.builder()
                        .rowCacheSize(1000)
                        .bufferSize(4096)
                        .open(is);
                System.out.println("FOI");
            }else{
                return new WorkbookHelper(file).wb;
            }
            return myWorkBook;
        }catch(IOException e){
            throw new BotCommandException("Error: " + e.getMessage());
        }
    }
}
