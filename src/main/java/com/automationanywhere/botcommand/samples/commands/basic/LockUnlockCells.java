package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.data.impl.TableValue;
import com.automationanywhere.botcommand.data.model.table.Row;
import com.automationanywhere.botcommand.data.model.table.Table;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.utils.FindInListSchema;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.GreaterThanEqualTo;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Spliterator;

import static com.automationanywhere.commandsdk.model.AttributeType.*;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

//BotCommand makes a class eligible for being considered as an action.
//@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "LockUnlockCells",
        label = "LockUnlockCells",
        node_label = "{{lock}} {{range}} cells from {{file}}",
        description = "",
        icon = "pkg.svg"
)

public class LockUnlockCells {
    @Execute
    public void action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "Excel file",description = "example: C:\\folder\\file.xlsx")
            //@FileExtension("xlsx")
            @NotEmpty
            String file,

            @Idx(index = "2", type = SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "ByName", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "ByIndex", value = "index"))})
            @Pkg(label = "Sheet:", description = "", default_value = "name", default_value_type = DataType.STRING)
            @NotEmpty
            String getSheetBy,

            @Idx(index = "2.1.1", type = TEXT)
            @Pkg(label = "Insira o nome da sheet:",description = "Será adiconada no fim da tabela")
            @NotEmpty
            String sheetName,

            @Idx(index = "2.2.1", type = NUMBER)
            @Pkg(label = "Insira o index da sheet:")
            @NotEmpty
            @GreaterThanEqualTo(value = "0")
            Double sheetIndex,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "Insira as celulas desejadas:", description = "Exemplo: A:C ou A1:B2 ")
            @NotEmpty
            String range,

            @Idx(index = "4", type = SELECT, options = {
                    @Idx.Option(index = "4.1", pkg = @Pkg(label = "lock", value = "lock")),
                    @Idx.Option(index = "4.2", pkg = @Pkg(label = "unlock", value = "unlock"))})
            @Pkg(label = "lock/Unlock:", description = "", default_value = "name", default_value_type = DataType.STRING)
            @NotEmpty
                    String lock
) {
            //================================================================= CREATE WORKBOOK OBJECT
            WorkbookHelper wbH = new WorkbookHelper(file);

            //================================================================= GET SHEET
            Sheet mySheet = this.getSheet(getSheetBy,sheetName,sheetIndex,wbH);

            boolean lck = lock.equals("lock");
            List<CellStyle> LISTSTYLE = new ArrayList<>();
            List<CellStyle> LISTNEWSTYLE = new ArrayList<>();

            CellRangeAddress a = CellRangeAddress.valueOf(range);
            for(CellAddress d: a){
                System.out.println("STR:" + d.toString() + " C:" + d.getColumn() + " R:" + d.getRow());

                Cell cl = wbH.getCell(mySheet,d);
                if (cl != null){
                    if(!LISTSTYLE.contains(cl.getCellStyle())){
                        CellStyle cstl = wbH.wb.createCellStyle();
                        cstl.cloneStyleFrom(cl.getCellStyle());
                        cstl.setLocked(lck);

                        LISTSTYLE.add(cl.getCellStyle());
                        LISTNEWSTYLE.add(cstl);
                    }


                    //CellStyle cstl = wbH.wb.createCellStyle();
                    //cstl.cloneStyleFrom(cl.getCellStyle());
                    //cstl.setLocked(lck);
                    Integer idx = LISTSTYLE.indexOf(cl.getCellStyle());
                    cl.setCellStyle(LISTNEWSTYLE.get(idx));


                    //System.out.println(cl);
                }

            }
            wbH.save(file);


//            CellReference cr = new CellReference("B2");
//            Cell cell = mySheet.getRow(cr.getRow()).getCell(cr.getCol());
//
//            CellStyle CellStyle = cell.getCellStyle();
//            unlockedCellStyle.setLocked(lck);
//
//            cell.setCellStyle(unlockedCellStyle);
//
//            wbH.save(file);

    }

        private Sheet getSheet(String getSheetBy, String sheetName, Double sheetIndex, WorkbookHelper wbH){
                Sheet mySheet = null;
                if(getSheetBy.equals("name")){
                        if(wbH.sheetExists(sheetName)){
                                wbH.wb.getSheet(sheetName);
                                mySheet = wbH.wb.getSheet(sheetName);
                        }else{
                                throw new BotCommandException("Sheet '" + sheetName + "' not found!");
                        }
                }else {
                        if(wbH.sheetExists(sheetIndex.intValue())){
                                mySheet = wbH.wb.getSheetAt(sheetIndex.intValue());
                        }else{
                                throw new BotCommandException("Sheet index '" + sheetIndex.intValue() + "' not found!");
                        }
                }
                return mySheet;
        }

}
