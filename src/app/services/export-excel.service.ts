import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { title } from 'process';
import * as logo from  './logo.js';

@Injectable({
  providedIn: 'root'
})
export class ExportExcelService {


  constructor() { }

  note1Excel() {


    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('note 1');
    worksheet.mergeCells('A1', 'F1');
    worksheet.mergeCells('A2', 'F2');
    worksheet.mergeCells('A3', 'F3');
    
    worksheet.mergeCells('A4', 'A6');
    worksheet.mergeCells('B4','B6');
    worksheet.mergeCells('D5','D6');
    worksheet.mergeCells('E5','E6');
    worksheet.mergeCells('C4','C6')
    worksheet.mergeCells('D4','F4');







    //worksheet.autoFileter = 'A1:G1';


    worksheet.getRow(1).height = 40;
    worksheet.getRow(2).heigth = 5;
    worksheet.getRow(3).height = 120;
    workbook.getRow(5).border 
    


    worksheet.columns = [
      {width: 30},
      {width: 7},
      {width: 30},
      {width: 15},
      {width: 15},
      {width: 15},


    ];

    // title row 

    let titleRow = worksheet.getCell('C1');
    
    titleRow.fill  = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'CDCDCD' }
    };

    titleRow.value = 'NOTE 1 : DETTES GARANTIES PAR DES SURETES REELLES';
    titleRow.font = {
      name: 'Calibri',
      size: 14,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' },
      background: {argb: '0085A3'}
    }
    titleRow.alignment = { vertical: 'middle', horizontal: 'center' }

    

    // end title row 

   
    let A4 = worksheet.getCell('A4');
    A4.value = 'LIBELES';
    A4.alignment = { vertical: 'middle', horizontal: 'center' }

    let B4 = worksheet.getCell('B4');
    B4.value = 'Note';
    B4.alignment = { vertical: 'middle', horizontal: 'center' }



    let C4 = worksheet.getCell('C4');
    C4.value = 'Montant brut';
    C4.alignment = { vertical: 'middle', horizontal: 'center' }

    let D4 = worksheet.getCell('D4')
    D4.value = 'SURETES REELLES'
    D4.alignment = { vertical: 'middle', horizontal: 'center' }


    let D5 = worksheet.getCell('D5')
    D5.value = 'Hypothèques'
    D5.alignment = { vertical: 'middle', horizontal: 'center' }

    let E5 = worksheet.getCell('E5')
    E5.value = 'Nantissements'
    E5.alignment = { vertical: 'middle', horizontal: 'center' }

    let F5 = worksheet.getCell('F5')
    F5.value = 'Gages/'
    F5.alignment = { vertical: 'middle', horizontal: 'center' }

    let F6 = worksheet.getCell('F6')
    F6.value = 'autres'
    F6.alignment = { vertical: 'middle', horizontal: 'center' }



    // row header 
    let headerRow = worksheet.getCell('C3');

    headerRow.value = `Désignation entité :………………………………………        Exercice  clos le 31-12-……………
    Numéro d’identification :………………………………      Durée (en mois) : ……………
    NOTE 1
    DETTES GARANTIES PAR DES SURETES REELLES`
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' }

    // end row header

    let r4 = worksheet.getCell('C4');
    r4.border = {
     // bottom: {style:'thin', color: {argb:'FFFFFF'}},
     
    };
    
    
    let r5 = worksheet.getCell('C5');
    r5.border = {
     // bottom: {style:'thin', color: {argb:'FFFFFF'}},
    };
    let r6 = worksheet.getCell('C6');

    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob,  'exported.xlsx');
    })

    


    /*

    //Title, Header & Data
    const title = excelData.title;
    const header = excelData.headers
    const data = excelData.data;

    //Create a workbook with a worksheet
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('Sales Data',{
      wiews: [{showGridlines: false}],
      pageSetup:{paperSize: 9, orientation:'portrait', fitToPage: true, fitToHeight: 5, fitToWidth: 7}
    });
    

    worksheet.autoFilter = "A4:C1"

    //Add Row and formatting
    worksheet.mergeCells('C1', 'F4');
    let titleRow = worksheet.getCell('C1');
    titleRow.value = title
    titleRow.font = {
      name: 'Calibri',
      size: 16,
      underline: 'single',

      bold: true,
      color: { argb: '0085A3' },
      background: {argb: '0085A3'}
    }
    titleRow.alignment = { vertical: 'middle', horizontal: 'center' }

    // Date
    worksheet.mergeCells('G1:H4');
    let d = new Date();
    let date = d.getDate() + '-' + d.getMonth() + '-' + d.getFullYear();
    let dateCell = worksheet.getCell('G1');
    dateCell.value = date;
    dateCell.font = {
      name: 'Calibri',
      size: 12,
      bold: true
    }
    dateCell.alignment = { vertical: 'middle', horizontal: 'center' }

    //Add Image
    let myLogoImage = workbook.addImage({
      base64: logo.logoBase64,
      extension: 'png',
    });
    worksheet.mergeCells('A1:B4');
    worksheet.addImage(myLogoImage, 'A1:B4');

    //Blank Row 
    worksheet.addRow([]);

    //Adding Header Row
    let headerRow = worksheet.addRow(header);
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: '4167B8' },
        bgColor: { argb: '' }
      }
      cell.font = {
        bold: true,
        color: { argb: 'FFFFFF' },
        size: 12
      }
    })

    // Adding Data with Conditional Formatting
    data.forEach(d => {
      let row = worksheet.addRow(d);

      let sales = row.getCell(6);
      let color = 'FF99FF99';
      if (+sales.value < 200000) {
        color = 'FF9999'
      }

      sales.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color }
      }
    }
    );

    worksheet.getColumn(3).width = 20;
    worksheet.addRow([]);

    //Footer Row
    let footerRow = worksheet.addRow(['Employee Sales Report Generated from example.com at ' + date]);
    footerRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFB050' }
    };


    worksheet.columns.forEach((col) => {
      col.style.font = { name: 'Comic Sans MS' };
      col.style.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  })

    //Merge Cells
    worksheet.mergeCells(`A${footerRow.number}:F${footerRow.number}`);

    //Generate & Save Excel File
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, title + '.xlsx');
    })

    */

  }


  exportEtatFinancier(){

    let workbook = new Workbook();
    let worksheetPG = workbook.addWorksheet('PAGE DE GARDE',{views: [{showGridLines: false}]});

    worksheetPG.getColumn(2).border = {
      bottom: {style:'thin', color: {argb:'000000'}},
    }
    
    worksheetPG.columns = [
      {width: 5},
      {width: 3},
      {width: 30},
      {width: 2},
      {width: 5},
      {width: 5},
      {width: 2},
      {width: 8},
      {width: 8},
      {width: 8},
      {width: 8},
      {width: 8},
      {width: 4},];


    worksheetPG.getRow(3).height = 30;
    worksheetPG.getCell('B4').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('B5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('C5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('D5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('E5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('F5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('G5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('H5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('I5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('J5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('K5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('L5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M5').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M4').border = {bottom: {style:'hair', color: {argb:'000000'}}}

    worksheetPG.getCell('B49').border = {top: {style:'hair', color: {argb:'000000'}}}

    worksheetPG.getCell('B52').border = {top: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('C51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('D51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('E51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('F51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('G51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('H51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('I51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('J51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('K51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('L51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M51').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M52').border = {top: {style:'hair', color: {argb:'000000'}}}

    worksheetPG.getCell('C12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('D12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('E12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('F12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('G12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('H12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('I12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('J12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('K12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('L12').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M12').border = {bottom: {style:'hair', color: {argb:'000000'}}}

    worksheetPG.getCell('C24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('D24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('E24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('F24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('G24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('H24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('I24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('J24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('K24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('L24').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M24').border = {bottom: {style:'hair', color: {argb:'000000'}}}


    worksheetPG.getCell('C26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('D26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('E26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('F26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('G26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('H26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('I26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('J26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('K26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('L26').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M26').border = {bottom: {style:'hair', color: {argb:'000000'}}}

    worksheetPG.getCell('C28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('D28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('E28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('F28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('G28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('H28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('I28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('J28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('K28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('L28').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M28').border = {bottom: {style:'hair', color: {argb:'000000'}}}


    worksheetPG.getCell('C30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('D30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('E30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('F30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('G30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('H30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('I30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('J30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('K30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('L30').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M30').border = {bottom: {style:'hair', color: {argb:'000000'}}}

    worksheetPG.getCell('C32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('D32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('E32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('F32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('G32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('H32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('I32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('J32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('K32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('L32').border = {bottom: {style:'hair', color: {argb:'000000'}}}
    worksheetPG.getCell('M32').border = {bottom: {style:'hair', color: {argb:'000000'}}}

    






    for(let i = 5; i < 52; i++){
      worksheetPG.getCell(`B${i}`).border = {left: {style:'hair', color: {argb:'000000'}}}
    }
    
    for(let i = 5; i < 52; i++){
      worksheetPG.getCell(`M${i}`).border = {right: {style:'hair', color: {argb:'000000'}}}
    }

    for(let i = 39; i <= 48; i++){
      worksheetPG.getCell(`C${i}`).border = {right: {style:'hair', color: {argb:'000000'}}, left: {style:'hair', color: {argb:'000000'}}}}
    
    





    worksheetPG.mergeCells('B3','M3');
    worksheetPG.mergeCells('F19','L19');
    worksheetPG.mergeCells('F37','L37');
   
    worksheetPG.mergeCells('H48','L48');
    worksheetPG.mergeCells('H44','L44');
    worksheetPG.mergeCells('H49','L50');
    worksheetPG.mergeCells('H40','L43');
    worksheetPG.mergeCells('H45','L47');
    worksheetPG.mergeCells('H39','L39');

   








    
    let B3 = worksheetPG.getCell('B3');

    B3.fill  = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '03BCF7' },
    };
    B3.font = {
      name: 'Calibri',
      size: 14,
      //underline: 'single',
      bold: true,
      color: { argb: 'FFFFFF' }
    }

    B3.value = "COVER SHEET";
    B3.alignment = { vertical: 'middle', horizontal: 'center' }



    let C5 = worksheetPG.getCell('C5');
    C5.value = "REPUBLIC OF SENEGAL";
    C5.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C5.alignment = { vertical: 'middle', horizontal: 'left' }


    let C7 = worksheetPG.getCell('C7');
    C7.value = "MINISTRY OF ECONOMY AND FINANCE";
    C7.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C7.alignment = { vertical: 'middle', horizontal: 'left' }


    let C9 = worksheetPG.getCell('C9');
    C9.value = "NATIONAL AUTHORITY OF TAXES AND PROPERTIES";
    C9.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C9.alignment = { vertical: 'middle', horizontal: 'left' }


    let C12 = worksheetPG.getCell('C12');
    C12.value = "FILING CENTER  : CENTRE DES GRANDES ENTREPRISES";
    C12.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C12.alignment = { vertical: 'middle', horizontal: 'left' }


    let F15 = worksheetPG.getCell('F15');
    F15.value = "STANDARDIZED FINANCIAL STATEMENTS OF THE ";
    F15.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    F15.alignment = { vertical: 'middle', horizontal: 'left' }

    let F16 = worksheetPG.getCell('F16');
    F16.value = "OHADA ACCOUNTING SYSTEM (SYSCOHADA)";
    F16.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    F16.alignment = { vertical: 'middle', horizontal: 'left' }


    let C19 = worksheetPG.getCell('C19');
    C19.value = "YEAR ENDED :";
    C19.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C19.alignment = { vertical: 'middle', horizontal: 'right' }


    let H19 = worksheetPG.getCell('H19');
    H19.value = "31st December 2019";
    H19.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    H19.alignment = { vertical: 'middle', horizontal: 'center' }

    let G22 = worksheetPG.getCell('G22');
    G22.value = "COMPANY DETAILS";
    G22.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    G22.alignment = { vertical: 'middle', horizontal: 'left' }


    let C24 = worksheetPG.getCell('C24');
    C24.value = "CORPORATE NAME :";
    C24.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C24.alignment = { vertical: 'middle', horizontal: 'left' }

    let C25 = worksheetPG.getCell('C25');
    C25.value = "(or name and surnames of owner)";
    C25.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C25.alignment = { vertical: 'middle', horizontal: 'left' }

    let C28 = worksheetPG.getCell('C28');
    C28.value = "ABBREVIATED NAME: ";
    C28.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C28.alignment = { vertical: 'middle', horizontal: 'left' }

    let C30 = worksheetPG.getCell('C30');
    C30.value = "COMPANY ADDRESS : ";
    C30.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C30.alignment = { vertical: 'middle', horizontal: 'left' }


    let C32 = worksheetPG.getCell('C32');
    C32.value = "TAX ID No.                    :                                                ";
    C32.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C32.alignment = { vertical: 'middle', horizontal: 'left' }


    let G34 = worksheetPG.getCell('G34');
    G34.value = "STANDARD SYSTEM";
    G34.font = {
      name: 'Times New Roman',
      size: 12,
      underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    G34.alignment = { vertical: 'middle', horizontal: 'left' }


    let C37 = worksheetPG.getCell('C37');
    C37.value = "Filed documents";
    C37.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C37.alignment = { vertical: 'middle', horizontal: 'left' }


    let H37 = worksheetPG.getCell('H37');
    H37.value = "For Tax Department use only";
    H37.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    H37.alignment = { vertical: 'middle', horizontal: 'center' }


    let H39 = worksheetPG.getCell('H39');
    H39.value = "Filing date";
    H39.border = {
      bottom: {style:'thin', color: {argb:'FFFFFF'}},
    }
    H39.font = {
      name: 'Times New Roman',
      size: 11,
    
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    H39.border = {
      top: {style: "thin", color: {rgba: '000000'}},
      left: {style: "thin", color: {rgba: '000000'}},
      right: {style: "thin", color: {rgba: '000000'}},

    }
    H39.alignment = { vertical: 'middle', horizontal: 'center' }


    let H40 = worksheetPG.getCell('H40');
    H40.border = {
      top: {style: "thin", color: {rgba: '000000'}},
      left: {style: "thin", color: {rgba: '000000'}},
      right: {style: "thin", color: {rgba: '000000'}},
      bottom: {style: "thin", color: {rgba: '000000'}},

      

    }

    let H44 = worksheetPG.getCell('H44');
    H44.value = "Name of the tax agent acknowledging";
    H44.font = {
      name: 'Times New Roman',
      size: 11,
      //underline: 'single',
      bold: false,
      color: { argb: '000000' }
    }
    H44.border = {
      top: {style: "thin", color: {rgba: '000000'}},
      left: {style: "thin", color: {rgba: '000000'}},
      right: {style: "thin", color: {rgba: '000000'}},
      bottom: {style: "thin", color: {rgba: '000000'}},

      

    }
    H44.alignment = { vertical: 'middle', horizontal: 'center' }


    let H45 = worksheetPG.getCell('H45');
    H45.border = {
      top: {style: "thin", color: {rgba: '000000'}},
      left: {style: "thin", color: {rgba: '000000'}},
      right: {style: "thin", color: {rgba: '000000'}},
      bottom: {style: "thin", color: {rgba: '000000'}},

      

    }

    let H48 = worksheetPG.getCell('H48');
    H48.value = "Signature and  stamp of the Tax Department";
    H48.font = {
      name: 'Times New Roman',
      size: 11,
      //underline: 'single',
      bold: false,
      color: { argb: '000000' }
    }
    H48.border = {
      top: {style: "thin", color: {rgba: '000000'}},
      left: {style: "thin", color: {rgba: '000000'}},
      right: {style: "thin", color: {rgba: '000000'}},
      bottom: {style: "thin", color: {rgba: '000000'}},

      

    }
    H48.alignment = { vertical: 'middle', horizontal: 'center' }


    let H49 = worksheetPG.getCell('H49');
  
    H49.border = {
      top: {style: "thin", color: {rgba: '000000'}},
      left: {style: "thin", color: {rgba: '000000'}},
      right: {style: "thin", color: {rgba: '000000'}},
      bottom: {style: "thin", color: {rgba: '000000'}},
   }

    let C39 = worksheetPG.getCell('C39');
    C39.value = "Indentification and miscellaneous information form";
    C39.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C39.border = {
      top: {style: "thin", color: {rgba: '000000'}}
    }
    C39.alignment = { vertical: 'middle', horizontal: 'left' }

    let C41 = worksheetPG.getCell('C41');
    C41.value = "Statement of financial position";
    C41.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C41.alignment = { vertical: 'middle', horizontal: 'left' }


    let C43 = worksheetPG.getCell('C43');
    C43.value = "Statement of Profit and Loss";
    C43.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C43.alignment = { vertical: 'middle', horizontal: 'left' }


    let C45 = worksheetPG.getCell('C45');
    C45.value = "Statement of Cashflow";
    C45.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C45.alignment = { vertical: 'middle', horizontal: 'left' }


    let C46 = worksheetPG.getCell('C46');
    C46.value = "Notes to the financial statements";
    C46.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C46.alignment = { vertical: 'middle', horizontal: 'left' }

    let C48 = worksheetPG.getCell('C48');
    C48.value = "Number of filed pages per copy :";
    C48.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C48.alignment = { vertical: 'middle', horizontal: 'left' }


    let C50 = worksheetPG.getCell('C50');
    C50.value = "Number of filed copies : ";
    C50.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    C50.alignment = { vertical: 'middle', horizontal: 'left' }


    let E39 = worksheetPG.getCell('E39');
    E39.value = "X";
    E39.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    E39.alignment = { vertical: 'middle', horizontal: 'center' }

    let E41 = worksheetPG.getCell('E41');
    E41.value = "X";
    E41.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    E41.alignment = { vertical: 'middle', horizontal: 'center' }


    let E43 = worksheetPG.getCell('E43');
    E43.value = "X";
    E43.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    E43.alignment = { vertical: 'middle', horizontal: 'center' }


    let E45 = worksheetPG.getCell('E45');
    E45.value = "X";
    E45.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    E45.alignment = { vertical: 'middle', horizontal: 'center' }


    let E46 = worksheetPG.getCell('E46');
    E46.value = "X";
    E46.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    E46.alignment = { vertical: 'middle', horizontal: 'center' }

    let E48 = worksheetPG.getCell('E48');
    E48.value = "X";
    E48.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    E48.alignment = { vertical: 'middle', horizontal: 'center' }


    let E50 = worksheetPG.getCell('E50');
    E50.value = "X";
    E50.font = {
      name: 'Times New Roman',
      size: 12,
      //underline: 'single',
      bold: true,
      color: { argb: '000000' }
    }
    E50.alignment = { vertical: 'middle', horizontal: 'center' }







    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob,  'exported.xlsx');
    })




  }

  
}