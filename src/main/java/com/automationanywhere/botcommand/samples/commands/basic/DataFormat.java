package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.GreaterThanEqualTo;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

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

public class DataFormat {
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

            //================================================================ FORMATO
            @Idx(index = "4", type = CHECKBOX)
            @Pkg(label = "Formato",default_value = "false",default_value_type = DataType.BOOLEAN)
            @NotEmpty
                    Boolean formato,
            @Idx(index = "4.1", type = SELECT, options = {
                    @Idx.Option(index = "4.1.1", pkg = @Pkg(label = "Geral", value = "0")),
                    @Idx.Option(index = "4.1.2", pkg = @Pkg(label = "Numero", value = "2")),
                    @Idx.Option(index = "4.1.3", pkg = @Pkg(label = "Moeda", value = "165")),
                    @Idx.Option(index = "4.1.4", pkg = @Pkg(label = "Contabil", value = "44")),
                    @Idx.Option(index = "4.1.5", pkg = @Pkg(label = "Data Abreviada", value = "14")),
                    @Idx.Option(index = "4.1.6", pkg = @Pkg(label = "Data Completa", value = "166")),
                    @Idx.Option(index = "4.1.7", pkg = @Pkg(label = "Hora", value = "167")),
                    @Idx.Option(index = "4.1.8", pkg = @Pkg(label = "Porcentagem", value = "10")),
                    @Idx.Option(index = "4.1.9", pkg = @Pkg(label = "Fração", value = "12")),
                    @Idx.Option(index = "4.1.10", pkg = @Pkg(label = "Cientifico", value = "11")),
                    @Idx.Option(index = "4.1.11", pkg = @Pkg(label = "Texto", value = "49")),
                    @Idx.Option(index = "4.1.12", pkg = @Pkg(label = "Customizado", value = "custom"))
            })
            @Pkg(label = "Formatos:", description = "", default_value = "0", default_value_type = DataType.STRING)
            @NotEmpty
                    String formato_value,
            @Idx(index = "4.1.12.1", type = TEXT)
            @Pkg(label = "Customizado:",description = "$#,##0.00")
            @NotEmpty
                    String custom_value,
            //================================================================ ALINHAMENTO
            @Idx(index = "5", type = CHECKBOX)
            @Pkg(label = "Alinhamento",default_value = "false",default_value_type = DataType.BOOLEAN)
            @NotEmpty
                    Boolean alinhamento,
            @Idx(index = "5.1", type = SELECT, options = {
                    @Idx.Option(index = "5.1.1", pkg = @Pkg(label = "Esquerda", value = "LEFT")),
                    @Idx.Option(index = "5.1.2", pkg = @Pkg(label = "Centro", value = "CENTER")),
                    @Idx.Option(index = "5.1.3", pkg = @Pkg(label = "Direita", value = "RIGHT"))
            })
            @Pkg(label = "Alinhamentos:", description = "", default_value = "RIGHT", default_value_type = DataType.STRING)
            @NotEmpty
                    String alinhamento_value,
            //================================================================ ALINHAMENTO VERTICAL
            @Idx(index = "6", type = CHECKBOX)
            @Pkg(label = "Alinhamento Vertical",default_value = "false",default_value_type = DataType.BOOLEAN)
            @NotEmpty
                    Boolean alinhamentoVertical,
            @Idx(index = "6.1", type = SELECT, options = {
                    @Idx.Option(index = "6.1.1", pkg = @Pkg(label = "A cima", value = "TOP")),
                    @Idx.Option(index = "6.1.2", pkg = @Pkg(label = "Ao Centro", value = "CENTER")),
                    @Idx.Option(index = "6.1.3", pkg = @Pkg(label = "A Baixo", value = "BOTTOM"))
            })
            @Pkg(label = "Alinhamentos:", description = "", default_value = "BOTTOM", default_value_type = DataType.STRING)
            @NotEmpty
                    String alinhamentoVertical_value,
            //================================================================ BLOQUEIO DE CELULAS
            @Idx(index = "7", type = CHECKBOX)
            @Pkg(label = "Bloquear",default_value = "false",default_value_type = DataType.BOOLEAN)
            @NotEmpty
                    Boolean bloquear,
            @Idx(index = "7.1", type = BOOLEAN)
            @Pkg(label = "Bloquear:",default_value = "false",default_value_type = DataType.BOOLEAN)
            @NotEmpty
                    Boolean bloquear_value

) {
            //================================================================= CREATE WORKBOOK OBJECT
            WorkbookHelper wbH = new WorkbookHelper(file);

            //================================================================= GET SHEET
            Sheet mySheet = this.getSheet(getSheetBy,sheetName,sheetIndex,wbH);

            List<CellStyle> LISTSTYLE = new ArrayList<>();
            List<CellStyle> LISTNEWSTYLE = new ArrayList<>();

            CellRangeAddress a = CellRangeAddress.valueOf(range);
            for(CellAddress d: a){
                //Cell cl = wbH.getCell(mySheet,d);
                Cell cl = wbH.getCell(mySheet,d);

                if (cl != null) {
                    if(!LISTSTYLE.contains(cl.getCellStyle())){
                        CellStyle cstl = wbH.wb.createCellStyle();
                        cstl.cloneStyleFrom(cl.getCellStyle());

                        if (formato) {
                            if(!formato_value.equals("custom")) {
                                cstl.setDataFormat((short) Integer.parseInt(formato_value));
                            }else{
                                //org.apache.poi.ss.usermodel.DataFormat b =  wbH.wb.createDataFormat();
                                cstl.setDataFormat(wbH.wb.createDataFormat().getFormat(custom_value));
                                //CellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(custom_value));
                            }
                        }
                        if (alinhamento) {
                            cstl.setAlignment(HorizontalAlignment.valueOf(alinhamento_value));
                        }
                        if (alinhamentoVertical) {
                            cstl.setVerticalAlignment(VerticalAlignment.valueOf(alinhamentoVertical_value));
                        }
                        if (bloquear) {
                            cstl.setLocked(bloquear_value);
                        }

                        LISTSTYLE.add(cl.getCellStyle());
                        LISTNEWSTYLE.add(cstl);
                    }


                    Integer idx = LISTSTYLE.indexOf(cl.getCellStyle());
                    cl.setCellStyle(LISTNEWSTYLE.get(idx));
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
                System.out.println("FOI:" + mb);
            }else{
                return new WorkbookHelper(file).wb;
            }
            return myWorkBook;
        }catch(IOException e){
            throw new BotCommandException("Error: " + e.getMessage());
        }
    }
}
