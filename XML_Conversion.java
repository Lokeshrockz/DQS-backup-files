package com.exampleexcel.demoexcelread;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

public class XML_Conversion {
  public static void main(String[] args) throws IOException, ParserConfigurationException, TransformerException {

    Workbook workbook = new XSSFWorkbook("/Users/logeshpandij/Downloads/TP_DynamicQuota22 (1).xlsx");
    Sheet sheet = workbook.getSheetAt(0);

    DocumentBuilderFactory documentFactory = DocumentBuilderFactory.newInstance();
    DocumentBuilder documentBuilder = documentFactory.newDocumentBuilder();
    Document document = documentBuilder.newDocument();
    document.setXmlStandalone(true);
        
  //   public void validateDqs(String filePath) {
  //     System.out.println("Dynamic Quota Selector Data Validation Starts......");
  //     try {
  //         if (!this.sheet.isEmpty()) {
  //             this.dqsDf["Validation"] = this.dqsDf["ruleName"].isBlank() ? "Invalid Rule Name" : "Valid Rule Name";
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["ruleOrder"] == 0 ? "Invalid Rule Order" : "Valid Rule Order");
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["paygResource"] == 90006666 ? "Valid PayG Resource" : "Invalid PayG Resource");
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["paygValue"] == -1 || this.dqsDf["paygValue"] == 0 ? "Valid PayG Value" : "Invalid PayG Value");
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["ratType"] >= 0 ? "Valid RAT Type" : "Invalid RAT Type");
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["rgType"] >= 0 ? "Valid RG Type" : "Invalid RG Type");
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["resourceId"] >= 0 ? "Valid Resource Id" : "Invalid Resource Id");
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["dqFrom"] >= 0 ? "Valid From" : "Invalid From");
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["To"] >= 0 ? "Valid To" : "Invalid To");
  //             this.dqsDf["Validation"] += "," + (this.dqsDf["grantQuota"] == 0 ? "Invalid Grant Quota" : "Valid Grant Quota");
  //             this.dqsDf["Validation"] += "," + (cfg.unit.contains(this.dqsDf["grantUnit"]) ? "Valid Grant Unit" : "Invalid Grant Unit");
  //             System.out.println("Dynamic Quota Selector Data Validation Ends......");
  //             String validationFilePath = cfg.VALIDATION_FILE_PATH + "validation_" +
  //                     FilenameUtils.getBaseName(filePath) +
  //                     FilenameUtils.getExtension(filePath);
  //             if (this.dqsDf["Validation"].str.contains("Invalid").any()) {
  //                 this.dqsDf["Validation"] = this.dqsDf["Validation"].apply(this::getInvalidMsg);
  //                 this.dqsDf = this.dqsDf.loc[:, :'Validation'];
  //                 if (new File(validationFilePath).exists()) {
  //                     new File(validationFilePath).delete();
  //                 }
  //                 try (ExcelWriter dqsDfToExcel = new ExcelWriter(validationFilePath)) {
  //                     this.dqsDf.toExcel(dqsDfToExcel);
  //                     dqsDfToExcel.save();
  //                 }
  //                 System.out.println("Validation Data is written to Validation Excel File successfully.");
  //             } else {
  //                 System.out.println("Dynamic Quota Selector Data Validation Success......");
  //                 if (new File(validationFilePath).exists()) {
  //                     new File(validationFilePath).delete();
  //                 }
  //                 this.generateDqs();
  //             }
  //         }
  //     } catch (Exception e) {
  //         // Handle the exception
  //     }
  // }
  
    
    Element root = document.createElement("cim:ConfigObjects"); //%r
    root.setAttribute("xmlns:cim", "http://xmlns.oracle.com/communications/platform/model/Config");
    document.appendChild(root);

    Element dqs = document.createElement("dynamicQuotaSelectors"); 
    dqs.setAttribute("xmlns:mtd", "http://xmlns.oracle.com/communications/platform/model/Metadata");
    dqs.setAttribute("xmlns:pdc", "http://xmlns.oracle.com/communications/platform/model/pricing");
    root.appendChild(dqs);

    Element name = document.createElement("name"); 
    dqs.appendChild(name);
    name.appendChild( document.createTextNode("TP dynamic quota configuration"));

    Element dscpt = document.createElement("description"); 
    dqs.appendChild(dscpt);
    dscpt.appendChild( document.createTextNode("Dynamic Quota Configuration_Description"));

    Element prce = document.createElement("priceListName"); 
    dqs.appendChild(prce);
    prce.appendChild( document.createTextNode("Default"));

    Element applto = document.createElement("applicableToName"); 
    dqs.appendChild(applto);
    applto.appendChild( document.createTextNode("TelcoGsm"));

    Element evesp = document.createElement("eventSpecName"); 
    dqs.appendChild(evesp);
    evesp.appendChild( document.createTextNode("EventDelayedSessionTelcoGprs"));

    Element chieve = document.createElement("applicableToAllChildEvent"); 
    dqs.appendChild(chieve);
    chieve.appendChild( document.createTextNode("false"));

    Element ValPer= document.createElement("validityPeriod"); 
    dqs.appendChild(ValPer);

    Element Valfro= document.createElement("validFrom"); 
    ValPer.appendChild(Valfro);
    Valfro.appendChild( document.createTextNode("20170921"));

 
    ArrayList<String> list=new ArrayList<String>();
    // System.out.println(sheet.getLastRowNum());
    for (int i = 1 ; i <= 22 ; i++) {
     
     System.out.println(">>>>>>>>>>>>>>>>>"+sheet.getLastRowNum() );
      Row row = sheet.getRow(i);

      
      Element rowElement = document.createElement("rule");
      ValPer.appendChild(rowElement);

      if (i == 0) {
        for (int j = 1; j <= row.getLastCellNum(); j++) {
          
          list.add(String.valueOf(row.getCell(j)).replace(".", "").replace(" ", "")+""); 
          System.out.println(list.get(j));
        }
      } 
      // else {
        for (int j = 1; j <=row.getLastCellNum(); j++) {

          if (j <= list.size() ) { 
            // System.out.println(list.get(j));
            // System.out.println(String.valueOf(row.getCell(j)));
            Element cellElement = document.createElement(list.get(j));
            if(j==1){
              list.get(j);
            }
            cellElement.appendChild(document.createTextNode(row.getCell(j).toString()));
            rowElement.appendChild(cellElement);
            
          }
        }
      // }


    String inputFilePath = "/Users/logeshpandij/Downloads/TP_DynamicQuota.xlsx";
      File inputFile = new File(inputFilePath);
      DataFormatter dataFormatter = new DataFormatter();
        
      if (inputFile.exists() && !inputFile.isDirectory()) {
          // DynamicQuotaSelector objDqs = new DynamicQuotaSelector();
          // objDqs.readDqs(inputFile);
          System.out.println("File exists");

          for (Row row1 : sheet) {
            // Skip the header row
            if (row1.getRowNum() == 1) {
                continue;
            } 
            if(row1.getRowNum()> sheet.getLastRowNum()){
              System.exit(0);
            }
              
            // Validate the rule name
            Cell ruleNameCell = row.getCell(1);
            String ruleName = dataFormatter.formatCellValue(ruleNameCell);
            if (ruleName.isBlank()) {
                System.out.println("Invalid Rule Name at row " + (row.getRowNum() + 1));
            }

            // Validate the rule order
            Cell ruleOrderCell = row.getCell(2);
            if(ruleOrderCell != null) {
              if(ruleOrderCell.getCellType() == CellType.NUMERIC){
             int ruleOrder = (int) ruleOrderCell.getNumericCellValue();
          }  else {
            System.out.println("Invalid Rule Order at row " + (row.getRowNum() + 1));
            }
          }else {
              System.out.println("Invalid Rule Order at row " + (row.getRowNum() + 1));
              
          }

            // Validate the payg resource
            Cell paygResourceCell = row.getCell(3);
if (paygResourceCell.getCellType() == CellType.NUMERIC) {
    int paygResource = (int) paygResourceCell.getNumericCellValue();
    if (paygResource != 90006666) {
        System.out.println("Invalid PayG Resource at row " + (row.getRowNum() + 1));
    }
} else {
    System.out.println("PayG Resource is not numeric at row " + (row.getRowNum() + 1));
}


            // Validate the payg value
            Cell paygValueCell = row.getCell(4);
            int paygValue = (int) paygValueCell.getNumericCellValue();
            if (paygValue != -1 && paygValue != 0) {
            System.out.println("Invalid PayG Value at row " + (row.getRowNum() + 1));
            }

            // Validate the rat type
            Cell ratTypeCell = row.getCell(5);
            int ratType = (int) ratTypeCell.getNumericCellValue();
            if (ratType < 0) {
            System.out.println("Invalid RAT Type at row " + (row.getRowNum() + 1));
            }

            // Validate the rg type
            Cell rgTypeCell = row.getCell(6);
            int rgType;
          if (rgTypeCell.getCellType() == CellType.NUMERIC) {
              rgType = (int) rgTypeCell.getNumericCellValue();
            } else {
              String rgTypeString = rgTypeCell.getStringCellValue();
                if (rgTypeString.equalsIgnoreCase("NA")) {
                rgType = 0;
              } else {
                System.out.println("Invalid RG Type at row " + (row.getRowNum() + 1));
                }
              }

            // Validate the resource id
            Cell resourceIdCell = row.getCell(7);
            int resourceId = (int) resourceIdCell.getNumericCellValue();
            if (resourceId < 0) {
            System.out.println("Invalid Resource Id at row " + (row.getRowNum() + 1));
            }

             // Validate the from value
            Cell fromCell = row.getCell(8);
            int from = (int) fromCell.getNumericCellValue();
            if (from < 0) {
            System.out.println("Invalid From at row " + (row.getRowNum() + 1));
            }

             // Validate the to value
            Cell toCell = row.getCell(9);
            int to = (int) toCell.getNumericCellValue();
          if (to < 0) {
            System.out.println("Invalid To at row " + (row.getRowNum() + 1));
            }      

            // Validate the grant Quota
            Cell grantQuotaCell = row.getCell(10);
            int grantQuota = (int) grantQuotaCell.getNumericCellValue();
            if (grantQuota == 0 ){
              System.out.println("Invalid grant Quota at row " + (row.getRowNum() + 1));
            }

            Cell grantUnitCell = row.getCell(11);
            String grantUnit = dataFormatter.formatCellValue(grantUnitCell);
            if (!grantUnit.equals("SECONDS") || !grantUnit.equals("MINUTES") || !grantUnit.equals("HOURS") ||
              !grantUnit.equals("DAYS") || !grantUnit.equals("BYTES") || !grantUnit.equals("KBYTES") ||
              !grantUnit.equals("MBYTES") || !grantUnit.equals("GBYTES") || !grantUnit.equals("NONE")) {
                System.out.println("Invalid grant Unit at row " + (row.getRowNum() + 1));

              }
      }
    }
      
      else {
          System.out.println("Input file does not exist!");
      }
    

      Element rulnam = document.createElement("ruleName");     //ruleName
      rowElement.appendChild(rulnam);
      // rulnam.appendChild( document.createTextNode(row.getCell(1).getStringCellValue()));
      System.out.println("Processing row " + row.getRowNum() + ", cell " + 1);
        Cell cell = row.getCell(1);
        if (Objects.isNull(cell)) {
        rulnam.appendChild(document.createTextNode(""));
          } else {
        rulnam.appendChild(document.createTextNode(String.valueOf(cell)));
          }
  
// if (rulnam != null) {
//     rulnam.appendChild(document.createTextNode(row.getCell(1).getStringCellValue()));
// }
      Element rulordr = document.createElement("ruleOrder");    //ruleOrder
      rowElement.appendChild(rulordr);
      rulordr.appendChild(document.createTextNode(String.valueOf((int) row.getCell(2).getNumericCellValue())));
      // rulordr.appendChild(document.createTextNode(String.valueOf(int) row.getCell(2).getNumericCellValue()));

      Element dqfve = document.createElement("dynamicQuotaFieldToValueExpression"); 
      Element oprt = document.createElement("operator"); 
      Element fldname = document.createElement("fieldName"); 
      Element fldkd = document.createElement("fieldKind"); 
      Element fldval = document.createElement("fieldValue");                          

      // this if func works only if ratType > 0:
      if(row.getCell(5).getCellType() == CellType.NUMERIC && row.getCell(5).getNumericCellValue()  > 0) {
        //ratType value
        rowElement.appendChild(dqfve);
        
        dqfve.appendChild(oprt);
        oprt.appendChild( document.createTextNode("EQUAL_TO"));
  
        dqfve.appendChild(fldname);
        fldname.appendChild( document.createTextNode("EventDelayedSessionTelcoGprs.TELCO_INFO.PRIMARY_MSID"));
        
        dqfve.appendChild(fldkd);
        fldkd.appendChild( document.createTextNode("EVENT_SPEC_FIELD"));

        dqfve.appendChild(fldval);
        // fldval.appendChild(document.createTextNode(row.getCell(6).toString()));
        fldval.appendChild(document.createTextNode(String.valueOf((int) row.getCell(5).getNumericCellValue())));
      }

      // this if func proceeds only when the rgtype value >0 :

      if (row.getCell(6).getCellType() == CellType.NUMERIC && row.getCell(6).getNumericCellValue() > 0){
          //rgType value
        rowElement.appendChild(dqfve);
        
        dqfve.appendChild(oprt);
        oprt.appendChild( document.createTextNode("EQUAL_TO"));
  
        dqfve.appendChild(fldname);
        fldname.appendChild( document.createTextNode("EventDelayedSessionTelcoGprs.TELCO_INFO.PRIMARY_MSID"));
        
        dqfve.appendChild(fldkd);
        fldkd.appendChild( document.createTextNode("EVENT_SPEC_FIELD"));

        dqfve.appendChild(fldval);
        // fldval.appendChild(document.createTextNode(row.getCell(6).toString()));
        fldval.appendChild(document.createTextNode(String.valueOf((int) row.getCell(6).getNumericCellValue())));

    }

     Element dqce = document.createElement("dynamicQuotaComplexExpression"); 
     Element oprt1 = document.createElement("operator"); 
     Element val = document.createElement("value");   //paygValue
     Element dqbe = document.createElement("dynamicQuotaBinaryExpression"); 
     Element leop = document.createElement("leftOperand"); 
     Element dqne = document.createElement("dynamicQuotaNumberExpression"); 
     Element nmbr = document.createElement("number"); 
     Element riop = document.createElement("rightOperand"); 

    // this if func works when the paygResource and paygvalues as follows:

    if (row.getCell(3).getCellType() == CellType.NUMERIC && row.getCell(3).getNumericCellValue() == 90006666
      && (row.getCell(4).getCellType() == CellType.NUMERIC && row.getCell(4).getNumericCellValue() == -1
      || row.getCell(4).getCellType() == CellType.NUMERIC && row.getCell(4).getNumericCellValue() == 0)){

        rowElement.appendChild(dqce);

        dqce.appendChild(oprt1);
        oprt1.appendChild( document.createTextNode("EQUAL_TO"));

        dqce.appendChild(val);
        val.appendChild(document.createTextNode(String.valueOf((int) row.getCell(4).getNumericCellValue())));  
                                                                                                     //paygValue
        dqce.appendChild(dqbe);

        dqbe.appendChild(leop);

        leop.appendChild(dqne);

        dqne.appendChild(nmbr);
        nmbr.appendChild( document.createTextNode("1"));

        dqbe.appendChild(riop);

        Element dqbo = document.createElement("dynamicQuotaBinaryOperator"); 
        dqbe.appendChild(dqbo);
        dqbo.appendChild( document.createTextNode("MULTIPLY"));

        Element dqble = document.createElement("dynamicQuotaBalanceExpression"); 
        riop.appendChild(dqble);

        Element benc = document.createElement("balanceElementNumCode");   // paygResource
        dqble.appendChild(benc);
        // benc.appendChild(document.createTextNode("90006666"));
        // benc.appendChild( document.createTextNode(parseInt(row.getCell(3).toString())));
        benc.appendChild(document.createTextNode(String.valueOf((int) row.getCell(3).getNumericCellValue())));
        
      }
          //  This if condn works when resId >0 && From >0 && To == 0 :
          
      if (row.getCell(7).getCellType() == CellType.NUMERIC && row.getCell(7).getNumericCellValue() > 0) {
        System.out.println("Inside 1stloopp>>>");


            if( (row.getCell(8).getCellType() == CellType.NUMERIC &&  row.getCell(8).getNumericCellValue() > 0)
                 && (row.getCell(9).getCellType() == CellType.NUMERIC && row.getCell(9).getNumericCellValue() == 0)) 
                 {
                  System.out.println("Inside 2ndloopp>>>");
                    
                  Element resid_dqce = document.createElement("dynamicQuotaComplexExpression"); 
                  Element resid_oprt1 = document.createElement("operator"); 
                  Element resid_val = document.createElement("value");   //paygValue
                  Element resid_dqbe = document.createElement("dynamicQuotaBinaryExpression"); 
                  Element resid_leop = document.createElement("leftOperand"); 
                  Element resid_dqne = document.createElement("dynamicQuotaNumberExpression"); 
                  Element resid_nmbr = document.createElement("number"); 
                  Element resid_riop = document.createElement("rightOperand"); 

                  rowElement.appendChild(resid_dqce);

                  resid_dqce.appendChild(resid_oprt1);
                  resid_oprt1.appendChild( document.createTextNode("GREATER_THAN"));

                  resid_dqce.appendChild(resid_val);
                  resid_val.appendChild(document.createTextNode(String.valueOf((int) row.getCell(8).getNumericCellValue())));

                  resid_dqce.appendChild(resid_dqbe);

                  resid_dqbe.appendChild(resid_leop);

                  resid_leop.appendChild(resid_dqne);

                  resid_dqne.appendChild(resid_nmbr);
                  resid_nmbr.appendChild( document.createTextNode("-1"));

                  resid_dqbe.appendChild(resid_riop);

                  Element resid_dqbo = document.createElement("dynamicQuotaBinaryOperator"); 
                  resid_dqbe.appendChild(resid_dqbo);
                  resid_dqbo.appendChild( document.createTextNode("MULTIPLY"));

                  Element resid_dqble = document.createElement("dynamicQuotaBalanceExpression"); 
                  resid_riop.appendChild(resid_dqble);

                    Element resid_benc = document.createElement("balanceElementNumCode");   // paygResource
                    resid_dqble.appendChild(resid_benc);
                    // benc.appendChild(document.createTextNode("90006666"));
                      // benc.appendChild( document.createTextNode(parseInt(row.getCell(3).toString())));
                      resid_benc.appendChild(document.createTextNode(String.valueOf((int) row.getCell(7).getNumericCellValue())));

                    System.out.println( "1 if loop");
       }
                //  This if condn works when resId >0 && From == 0 && To > 0 :

        if(row.getCell(8).getCellType() == CellType.NUMERIC && row.getCell(8).getNumericCellValue() == 0
            && row.getCell(9).getCellType() == CellType.NUMERIC && row.getCell(9).getNumericCellValue() > 0) {
                    Element resid_dqce_lt = document.createElement("dynamicQuotaComplexExpression"); 
                    Element resid_oprt1_lt = document.createElement("operator"); 
                    Element resid_val_lt = document.createElement("value");   //paygValue
                    Element resid_dqbe_lt = document.createElement("dynamicQuotaBinaryExpression"); 
                    Element resid_leop_lt = document.createElement("leftOperand"); 
                    Element resid_dqne_lt = document.createElement("dynamicQuotaNumberExpression"); 
                    Element resid_nmbr_lt = document.createElement("number"); 
                    Element resid_riop_lt = document.createElement("rightOperand"); 

                    rowElement.appendChild(resid_dqce_lt);

                    resid_dqce_lt.appendChild(resid_oprt1_lt);
                    resid_oprt1_lt.appendChild( document.createTextNode("LESS_THAN"));

                    resid_dqce_lt.appendChild(resid_val_lt);
                    resid_val_lt.appendChild(document.createTextNode(String.valueOf((int) row.getCell(9).getNumericCellValue())));

                    resid_dqce_lt.appendChild(resid_dqbe_lt);

                    resid_dqbe_lt.appendChild(resid_leop_lt);

                    resid_leop_lt.appendChild(resid_dqne_lt);

                    resid_dqne_lt.appendChild(resid_nmbr_lt);
                    resid_nmbr_lt.appendChild( document.createTextNode("-1"));

                    resid_dqbe_lt.appendChild(resid_riop_lt);

                    Element resid_dqbo_lt = document.createElement("dynamicQuotaBinaryOperator"); 
                    resid_dqbe_lt.appendChild(resid_dqbo_lt);
                    resid_dqbo_lt.appendChild( document.createTextNode("MULTIPLY"));

                    Element resid_dqble_lt = document.createElement("dynamicQuotaBalanceExpression"); 
                    resid_riop_lt.appendChild(resid_dqble_lt);

                      Element resid_benc_lt = document.createElement("balanceElementNumCode");   // paygResource
                      resid_dqble_lt.appendChild(resid_benc_lt);
                      // benc.appendChild(document.createTextNode("90006666"));
                     // benc.appendChild( document.createTextNode(parseInt(row.getCell(3).toString())));
                     resid_benc_lt.appendChild(document.createTextNode(String.valueOf((int) row.getCell(7).getNumericCellValue())));
                     System.out.println( "2 if loop");

        }
              // this if func for from >0 and To < From  :
             if (row.getCell(9).getCellType() == CellType.NUMERIC && row.getCell(9).getCellType() == CellType.NUMERIC
             &&( row.getCell(8).getCellType() == CellType.NUMERIC && row.getCell(8).getNumericCellValue() > 0)
             && !(row.getCell(9).getNumericCellValue() < row.getCell(8).getNumericCellValue()))  {
                    //  System.out.println("eight:" + row.getCell(8) );

                    Element resid_dqce_gte = document.createElement("dynamicQuotaComplexExpression"); 
                    Element resid_oprt1_gte  = document.createElement("operator"); 
                    Element resid_val_gte  = document.createElement("value");   //paygValue
                    Element resid_dqbe_gte  = document.createElement("dynamicQuotaBinaryExpression"); 
                    Element resid_leop_gte  = document.createElement("leftOperand"); 
                    Element resid_dqne_gte  = document.createElement("dynamicQuotaNumberExpression"); 
                    Element resid_nmbr_gte  = document.createElement("number"); 
                    Element resid_riop_gte  = document.createElement("rightOperand"); 

                    rowElement.appendChild(resid_dqce_gte );

                    resid_dqce_gte .appendChild(resid_oprt1_gte );
                    resid_oprt1_gte .appendChild( document.createTextNode("GREATER_THAN_EQUAL"));

                    resid_dqce_gte .appendChild(resid_val_gte );
                    resid_val_gte .appendChild(document.createTextNode(String.valueOf((int) row.getCell(8).getNumericCellValue())));

                    resid_dqce_gte .appendChild(resid_dqbe_gte );

                    resid_dqbe_gte .appendChild(resid_leop_gte );

                    resid_leop_gte .appendChild(resid_dqne_gte );

                    resid_dqne_gte .appendChild(resid_nmbr_gte );
                    resid_nmbr_gte .appendChild( document.createTextNode("-1"));

                    resid_dqbe_gte .appendChild(resid_riop_gte );

                    Element resid_dqble_gte  = document.createElement("dynamicQuotaBalanceExpression"); 
                    resid_riop_gte .appendChild(resid_dqble_gte );

                    Element resid_benc_gte  = document.createElement("balanceElementNumCode");   // paygResource
                      resid_dqble_gte .appendChild(resid_benc_gte );
                      // benc.appendChild(document.createTextNode("90006666"));
                     // benc.appendChild( document.createTextNode(parseInt(row.getCell(3).toString())));
                     resid_benc_gte .appendChild(document.createTextNode(String.valueOf((int) row.getCell(7).getNumericCellValue())));

                    Element resid_dqbo_gte  = document.createElement("dynamicQuotaBinaryOperator"); 
                    resid_dqbe_gte .appendChild(resid_dqbo_gte );
                    resid_dqbo_gte .appendChild( document.createTextNode("MULTIPLY"));

                     System.out.println( "3 if loop");

                  // LESS THAN EQUAL
      
          Element resid_dqce2 = document.createElement("dynamicQuotaComplexExpression"); 
          Element resid_oprt2 = document.createElement("operator"); 
          Element resid_val2 = document.createElement("value");   //paygValue
          Element resid_dqbe2 = document.createElement("dynamicQuotaBinaryExpression"); 
          Element resid_leop2 = document.createElement("leftOperand"); 
          Element resid_dqne2 = document.createElement("dynamicQuotaNumberExpression"); 
          Element resid_nmbr2 = document.createElement("number"); 
          Element resid_riop2 = document.createElement("rightOperand"); 

          rowElement.appendChild(resid_dqce2);

          resid_dqce2.appendChild(resid_oprt2);
          resid_oprt2.appendChild( document.createTextNode("LESS_THAN_EQUAL"));

          resid_dqce2.appendChild(resid_val2);
          resid_val2.appendChild(document.createTextNode(String.valueOf((int) row.getCell(9).getNumericCellValue())));

          resid_dqce2.appendChild(resid_dqbe2);

          resid_dqbe2.appendChild(resid_leop2);

          resid_leop2.appendChild(resid_dqne2);

          resid_dqne2.appendChild(resid_nmbr2);
          resid_nmbr2.appendChild( document.createTextNode("-1"));

          resid_dqbe2.appendChild(resid_riop2);

          Element resid_dqbo2 = document.createElement("dynamicQuotaBinaryOperator"); 
          resid_dqbe2.appendChild(resid_dqbo2);
          resid_dqbo2.appendChild( document.createTextNode("MULTIPLY"));

          Element resid_dqble2 = document.createElement("dynamicQuotaBalanceExpression"); 
          resid_riop2.appendChild(resid_dqble2);

            Element resid_benc2 = document.createElement("balanceElementNumCode");   // paygResource
            resid_dqble2.appendChild(resid_benc2);
            // benc.appendChild(document.createTextNode("90006666"));
           // benc.appendChild( document.createTextNode(parseInt(row.getCell(3).toString())));
           resid_benc2.appendChild(document.createTextNode(String.valueOf((int) row.getCell(7).getNumericCellValue())));
       
    }
  }
        if(row.getCell(10).getCellType() == CellType.NUMERIC && row.getCell(10).getNumericCellValue()  > 0){

          Element config = document.createElement("configuration"); 
          rowElement.appendChild(config);

            Element key = document.createElement("key"); 
            config.appendChild(key);
          key.appendChild( document.createTextNode("VALIDITY_TIME"));      

          Element valu = document.createElement("value"); 
          config.appendChild(valu);
          valu.appendChild( document.createTextNode("100")); 

          Element unit = document.createElement("unit"); 
          config.appendChild(unit);
          unit.appendChild( document.createTextNode("MINUTES")); 

          Element config1 = document.createElement("configuration"); 
          rowElement.appendChild(config1);

          Element key1 = document.createElement("key"); 
          config1.appendChild(key1);
          key1.appendChild( document.createTextNode("QUOTA_HOLDING_TIME"));      

          Element valu1 = document.createElement("value"); 
          config1.appendChild(valu1);
          valu1.appendChild( document.createTextNode("20")); 

          Element unit1 = document.createElement("unit"); 
          config1.appendChild(unit1);
          unit1.appendChild( document.createTextNode("SECONDS")); 

          Element config2 = document.createElement("configuration"); 
          rowElement.appendChild(config2);

          Element key2 = document.createElement("key"); 
          config2.appendChild(key2);
          key2.appendChild( document.createTextNode("QUOTA_HOLDING_TIME"));      

          Element valu2 = document.createElement("value"); 
          config2.appendChild(valu2);
          valu2.appendChild( document.createTextNode("20")); 

          Element unit2 = document.createElement("unit"); 
          config2.appendChild(unit2);
          unit2.appendChild( document.createTextNode("SECONDS")); 

          Element requnit = document.createElement("requestedUnits");
          rowElement.appendChild(requnit);

          Element fldname1 = document.createElement("fieldName");
          requnit.appendChild(fldname1);
          fldname1.appendChild(document.createTextNode("EventDelayedSessionTelcoGprs.REQUESTED_UNITS.TOTAL_VOLUME"));

          Element unit3 = document.createElement("unit");
          requnit.appendChild(unit3);
          unit3.appendChild(document.createTextNode(row.getCell(11).toString()));

          Element dqnexp = document.createElement("dynamicQuotaNumberExpression");
          requnit.appendChild(dqnexp);

          Element nmbr1 = document.createElement("number");    // grantQuota 
          dqnexp.appendChild(nmbr1);
          nmbr1.appendChild(document.createTextNode(String.valueOf((int) row.getCell(10).getNumericCellValue())));  

      // nmbr1.appendChild(document.createTextNode(getCellValue(row.getCell(10))));

      // System.out.println();
      System.out.println( "5 if loop");
    }

}
     //  Element subElement = document.createElement("defaultRule");
      // root.appendChild(subElement);

      Element defrul = document.createElement("defaultRule"); 
      ValPer.appendChild(defrul);

      Element rulnam_def = document.createElement("ruleName");     //ruleName
      defrul.appendChild(rulnam_def);
      rulnam_def.appendChild( document.createTextNode("Rule_10000_PAYG_DEFAULT_QUOTA"));

       Element rulordr_def = document.createElement("ruleOrder");    //ruleOrder
          defrul.appendChild(rulordr_def);
        rulordr_def.appendChild( document.createTextNode("1000"));

        Element config3 = document.createElement("configuration"); 
         defrul.appendChild(config3);

          Element key3 = document.createElement("key"); 
          config3.appendChild(key3);
          key3.appendChild( document.createTextNode("VALIDITY_TIME"));      

          Element valu3 = document.createElement("value"); 
          config3.appendChild(valu3);
          valu3.appendChild( document.createTextNode("100")); 

          Element unit3 = document.createElement("unit"); 
          config3.appendChild(unit3);
          unit3.appendChild( document.createTextNode("MINUTES")); 

          Element config4 = document.createElement("configuration"); 
          defrul.appendChild(config4);

          Element key4 = document.createElement("key"); 
          config4.appendChild(key4);
          key4.appendChild( document.createTextNode("QUOTA_HOLDING_TIME"));      

          Element valu4 = document.createElement("value"); 
          config4.appendChild(valu4);
          valu4.appendChild( document.createTextNode("20")); 

          Element unit4 = document.createElement("unit"); 
          config4.appendChild(unit4);
          unit4.appendChild( document.createTextNode("SECONDS")); 

          Element config5 = document.createElement("configuration"); 
          defrul.appendChild(config5);

          Element key5 = document.createElement("key"); 
          config5.appendChild(key5);
          key5.appendChild( document.createTextNode("VOLUME_QUOTA_THRESHOLD"));      

          Element valu5 = document.createElement("value"); 
          config5.appendChild(valu5);
          valu5.appendChild( document.createTextNode("5")); 

          Element unit5 = document.createElement("unit"); 
          config5.appendChild(unit5);
          unit5.appendChild( document.createTextNode("KBYTES")); 

          Element requnit2 = document.createElement("requestedUnits");
          defrul.appendChild(requnit2);

          Element fldname2 = document.createElement("fieldName");
          requnit2.appendChild(fldname2);
          fldname2.appendChild(document.createTextNode("EventDelayedSessionTelcoGprs.REQUESTED_UNITS.TOTAL_VOLUME"));

          Element req_unit = document.createElement("unit");
          requnit2.appendChild(req_unit);
          req_unit.appendChild( document.createTextNode("KBYTES"));

          Element def_dqnexp = document.createElement("dynamicQuotaNumberExpression");
          requnit2.appendChild(def_dqnexp);

          Element nmbr2 = document.createElement("number");    // grantQuota 
          def_dqnexp.appendChild(nmbr2);
          nmbr2.appendChild( document.createTextNode("512"));

      TransformerFactory transformerFactory = TransformerFactory.newInstance();
      Transformer transformer = transformerFactory.newTransformer();

      transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
      transformer.setOutputProperty(OutputKeys.STANDALONE, "yes");
  
    DOMSource domSource = new DOMSource(document);
    StreamResult streamResult = new StreamResult(new FileOutputStream("newdemoww222.xml"));
    transformer.setOutputProperty(javax.xml.transform.OutputKeys.INDENT, "yes");
    transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", String.valueOf(3));

    transformer.transform(domSource, streamResult);

    System.out.println("XML file created successfully");
    workbook.close();
      }


  private static String parseInt(String string) {
    return null;
  }

  @Override
  public String toString() {
    return "XML_Conversion []";
  }

}

