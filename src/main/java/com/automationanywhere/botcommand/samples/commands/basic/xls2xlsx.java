package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.botcommand.data.model.table.Row;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.utils.Assets;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.GreaterThanEqualTo;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.script.Invocable;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static com.automationanywhere.commandsdk.model.AttributeType.*;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "DataFormat",
        label = "DataFormat",
        node_label = "set format {{range}} cells",
        description = "",
        icon = "pkg.svg"
)

public class xls2xlsx {
    @Execute
    public void action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Insira o nome da sheet:",description = "Será adiconada no fim da tabela")
            @NotEmpty
                    String from,
            @Idx(index = "2", type = TEXT)
            @Pkg(label = "Insira o nome da sheet:",description = "Será adiconada no fim da tabela")
            @NotEmpty
                    String to

) {
        //========================================================== LITURA DO JS
        ScriptEngineManager manager = new ScriptEngineManager();
        ScriptEngine engine = manager.getEngineByName("JavaScript");

        String library = new Assets().codeLibrary();

        try{
            //this.alert(Thread.currentThread().getContextClassLoader().getResource("./config/library.js").toString());
            //InputStream inputStream =  Thread.currentThread().getContextClassLoader().getResource("./config/library.js").openStream();
            //String library = new String(inputStream.readAllBytes(), StandardCharsets.UTF_8);
            engine.eval(library);

            //engine.eval(Files.newBufferedReader(Paths.get("C:/Temp/js.js"), StandardCharsets.UTF_8));
        }
        catch (Exception e){
            throw new BotCommandException("Error when trying to load Js code!" + e.getMessage());
        }

        //EXECUTA A FORMULA
        Invocable inv = (Invocable) engine;
        Object params[] = {from,to};

        try{
            Boolean filter = Boolean.parseBoolean(inv.invokeFunction("xls2xlsx", params).toString());
        }
        catch (Exception e){
            throw new BotCommandException("Error calling method 'xls2xlsx':" + e.getMessage());
        }




    }

}
