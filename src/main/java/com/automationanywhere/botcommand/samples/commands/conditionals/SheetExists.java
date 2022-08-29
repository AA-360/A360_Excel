/*
 * Copyright (c) 2020 Automation Anywhere.
 * All rights reserved.
 *
 * This software is the proprietary information of Automation Anywhere.
 * You shall use it only in accordance with the terms of the license agreement
 * you entered into with Automation Anywhere.
 */

/**
 *
 */
package com.automationanywhere.botcommand.samples.commands.conditionals;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.VariableType;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Workbook;
import org.ini4j.Ini;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static com.automationanywhere.commandsdk.annotations.BotCommand.CommandType.Condition;
import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

/**
 *
 * This example shows how to create an Condition.
 * <p>
 * Here we are checking of the provided boolean value is false.
 *
 *
 * @author Raj Singh Sisodia
 *
 */
@BotCommand(commandType = Condition)
@CommandPkg(
        label = "SheetExists",
        name = "SheetExists",
        description = "Se uma sheet existe",
        node_label = "Check if {{sheet}} exists in {{file}}"
)
public class SheetExists {

    @ConditionTest
    public Boolean validate(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Excel file",description = "example: C:\\folder\\file.xlsx")
            @NotEmpty
                    String file,
            @Idx(index = "2", type = AttributeType.SELECT, options = {
                    @Idx.Option(index ="2.1", pkg = @Pkg(label = "Equals(=)", value = "=")),
                    @Idx.Option(index ="2.2", pkg = @Pkg(label = "Diferent of (≠)", value = "≠")),
                    @Idx.Option(index ="2.3", pkg = @Pkg(label = "Includes", value = "in")),
                    @Idx.Option(index ="2.4", pkg = @Pkg(label = "Not Includes", value = "not in"))
            })
            @Pkg(label = "Operador",default_value_type = STRING,default_value = "=")
            @NotEmpty
                    String select,
            @Idx(index = "3", type = TEXT)
            @Pkg(label = "SheetName:")
            @NotEmpty
                    String value
    ) {
        Workbook myWorkBook = this.wbget(file);

        WorkbookHelper wbh = new WorkbookHelper(myWorkBook);
        List<String> lista = wbh.getSheetsName();

        System.out.println("FOI2");

        if(select.equals("=")){
            return lista.contains(value);
        }

        for(String sht: lista){
            if(select.equals("≠")){
                if(!sht.equals(value)) return true;
            }else if(select.equals("in")){
                if(sht.contains(value)) return true;
            }else{
                if(!sht.contains(value)) return true;
            }
        }
        return false;

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
