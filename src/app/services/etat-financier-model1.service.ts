import { Workbook } from 'exceljs';
import { Injectable } from '@angular/core';
import * as fs from 'file-saver';


@Injectable({
  providedIn: 'root'
})
export class EtatFinancierModel1Service {

  constructor() { }


  

  alphabet: string[] = [
     'A',
     'B',
     'C',
     'D',
     'E',
     'F',
     'G',
     'H',
     'I',
     'J',
     'K',
     'L',
     'M',
     'N',
     'O',
     'P',
     'Q',
     'R',
     'S',
     'T',
     'U',
     'V',
     'W',
     'X',
     'Y',
     'Z',
     'AA',
     'AB',
     'AC',
     'AD',
     'AE',
     'AF',
     'AG',
     'AH',
     'AI',
     'AJ',
     'AK',
     'AL',
     'AM'];


  exportmodel(){

    const workbook = new Workbook();
    const worksheetPg = workbook.addWorksheet('Page de garde',{views: [{showGridLines: false}]});


   const PGData = [
      {
        "lignes": 14
      },
      {
    "B3":"REPUBLIQUE",							
    "I3":"DU SENEGAL",													
    "B5":"MINISTERE",						
    "I5":"DES FINANCES ET DU BUDGET",													
    "B7":"DIRECTION",							
    "I7":"DIRECTION GENERALE DES IMÔTS ET DES DOMAINES",								
    "B15":"CENTRE DE DEPOT : DIRECTION DES GRANDES ENTREPRISES",																				
    "A21":"ETATS FINANCIERS NORMALISES DU SYSCOHADA",																					
    "B25":"EXERCICE CLOS LE :",										
    "M25":"31/12/2019",									
    "A28":"DESIGNATION DE L'ENTREPRISE",																				
    "C33":"DENOMINATION SOCIALE :",									
    "L33":"ENTREPRISE MORAL",									
    "C34":"(ou nom et prénoms de l'exploitant)",																			
    "C39":"SIGLE USUEL :",																		
    "C44":"ADRESSE COMPLETE :",							
    "J44":"ZONE 15",											
    "C49":"N° D'IDENTIFICATION FISCALE :",										
    "N49":"004942901",							
    "C53":"SYSTEME NORMAL",																			
    "C57":"Documents déposés",																		
    "V57":"Réservé à la Direction Générale  des Impôts",
    "C60":"Fiche d'identification et renseignements divers",																
    "S60":"X",	
    "W60":"Date de dépôt",
    "C62":"Bilan",															
    "S62":"X",			
    "C64":"Compte de résultat",															
    "S64":"X",	
    "W64":"Nom de l'agent de DGI ayant réceptionné le dépôt",
    "C66": "Tableau des flux de trésorie",
    "S66": "X",
    "C68":"Notes annexes",															
    "S68":"X",		
    "W68":"Signature de l'agent et cachet du service",
    "C70": "Nombre de pages déposées par exemplaire",
    "C71":"Nombre de pages déposées par exemplaire",
    "R71": 5															

      }
    ]

    worksheetPg.mergeCells('A15', 'AM15');
    worksheetPg.mergeCells('A21', 'AM21');
    worksheetPg.mergeCells('A22', 'AM22');
    worksheetPg.mergeCells('A53', 'AM53');
    worksheetPg.mergeCells('B25', 'K25');
    worksheetPg.mergeCells('M25', 'AK25');
    worksheetPg.mergeCells('A28', 'AM28');
    worksheetPg.mergeCells('L33', 'AK33');
    worksheetPg.mergeCells('J44', 'AK44');
    worksheetPg.mergeCells('N49', 'AJ49');
    worksheetPg.mergeCells('C57', 'T57');
    worksheetPg.mergeCells('V57', 'AK57');
    worksheetPg.mergeCells('W60', 'AK60');
    worksheetPg.mergeCells('W64', 'AK64');
    worksheetPg.mergeCells('W62', 'AK62');
    worksheetPg.mergeCells('W68', 'AK68');
    worksheetPg.mergeCells('R71', 'T71');
    worksheetPg.mergeCells('I3', 'Y3');
    worksheetPg.mergeCells('I5', 'Y5');
    worksheetPg.mergeCells('I7', 'Y7');





    worksheetPg.columns = [
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},
      {width: 4},];



      for(let i = 0; i <= Object.keys(PGData).length; i++ ){

        for(var v in PGData[1]){
    
        
             console.log(`mon ${v}`)
              let cell = worksheetPg.getCell(v);
             cell.value = PGData[1][v];
              cell.font = {
                name: 'Cambria',
                size: 9,
                //underline: 'single',
               // bold: true,
                color: { argb: '000000' }
              }
        
              cell.alignment = { vertical: 'middle', horizontal: 'justify'}
    
            
            
      
        }
    
      }






      for(let i = 1; i < 73; i++){
        worksheetPg.getCell(`AM${i}`).border = {right: {style:'hair', color: {argb:'000000'}}}
      }

      for(let i = 59; i < 71; i++){
        worksheetPg.getCell(`T${i}`).border = {right: {style:'hair', color: {argb:'000000'}}}
      }

      for(let i = 59; i < 71; i++){
        worksheetPg.getCell(`C${i}`).border = {
          left: {style:'hair', color: {argb:'000000'}},
          

        }
      }

      

      for(let i = 59; i < 70; i++){
        worksheetPg.getCell(`V${i}`).border = {
          left: {style:'hair', color: {argb:'000000'}},
        }
      }

      for(let i = 0; i < this.alphabet.length; i++){
        worksheetPg.getCell(`${this.alphabet[i]}73`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
        };
      }

      for(let i = 21; i < this.alphabet.length - 1; i++){
        worksheetPg.getCell(`${this.alphabet[i]}59`).border = {
          top: {style:'hair', color: {argb:'000000'}},
        };
      }

      for(let i = 22; i < this.alphabet.length - 2; i++){
        worksheetPg.getCell(`${this.alphabet[i]}71`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
        };
      }

      for(let i = 22; i < this.alphabet.length - 2; i++){
        worksheetPg.getCell(`${this.alphabet[i]}63`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
        };
      }

      for(let i = 22; i < this.alphabet.length - 2; i++){
        worksheetPg.getCell(`${this.alphabet[i]}67`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
        };
      }

      for(let i = 2; i < 20; i++){
        worksheetPg.getCell(`${this.alphabet[i]}69`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
          right: this.alphabet[i] === 'T' || this.alphabet[i] === 'Q' ? {style:'hair', color: {argb:'000000'}} : {},
          left: this.alphabet[i] === 'C' ? {style:'hair', color: {argb:'000000'}} : {},

          
        };
      }

      for(let i = 2; i < 20; i++){
        worksheetPg.getCell(`${this.alphabet[i]}58`).border = {bottom: {style:'hair', color: {argb:'000000'}}};
      }
      for(let i = 2; i < 20; i++){
        worksheetPg.getCell(`${this.alphabet[i]}71`).border = {
          bottom:  {style:'hair', color: {argb:'000000'}},
          right: this.alphabet[i] === 'T' ? {style:'hair', color: {argb:'000000'}} : {},
          left: this.alphabet[i] === 'C' ? {style:'hair', color: {argb:'000000'}} : {}

        };
      }

      for(let i = 59; i < 70; i++){
        worksheetPg.getCell(`R${i}`).border = {
          left: {style:'hair', color: {argb:'000000'}},
        }
      }

      for(let i = 59; i < 72; i++){
        worksheetPg.getCell(`V${i}`).border = {
          right: {style:'hair', color: {argb:'000000'}},
        }
      }

      for(let i = 59; i < 72; i++){
        worksheetPg.getCell(`AL${i}`).border = {
          left: {style:'hair', color: {argb:'000000'}},
        }
      }

      worksheetPg.getCell('R69').border = {
        bottom: {style:'hair', color: {argb:'000000'}},
      };

      

      let B3 = worksheetPg.getCell('B3');
      B3.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      B3.alignment = { vertical: 'middle', horizontal: 'left' }

      let B5 = worksheetPg.getCell('B5');
      B5.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      B5.alignment = { vertical: 'middle', horizontal: 'left' }

      let B7 = worksheetPg.getCell('B7');
      B7.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      B7.alignment = { vertical: 'middle', horizontal: 'left' }

      let I3 = worksheetPg.getCell('I3');
      I3.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      B3.alignment = { vertical: 'middle', horizontal: 'left' }

      let I5 = worksheetPg.getCell('I5');
      I5.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      I5.alignment = { vertical: 'middle', horizontal: 'left' }

      let I7 = worksheetPg.getCell('I7');
      I7.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      I7.alignment = { vertical: 'middle', horizontal: 'left' }

      let A15 = worksheetPg.getCell('A15');
      A15.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      A15.alignment = { vertical: 'middle', horizontal: 'center' }

      let A21 = worksheetPg.getCell('A21');
      A21.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      A21.alignment = { vertical: 'middle', horizontal: 'center' }

      let A22 = worksheetPg.getCell('A22');
      A22.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      A22.alignment = { vertical: 'middle', horizontal: 'center' }

      let B25 = worksheetPg.getCell('B25');
      B25.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      B25.alignment = { vertical: 'middle', horizontal: 'center' }

      let M25 = worksheetPg.getCell('M25');
      M25.value = "31/12/2019";
      M25.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      M25.alignment = { vertical: 'middle', horizontal: 'center' }

      let A28 = worksheetPg.getCell('A28');
      A28.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      A28.alignment = { vertical: 'middle', horizontal: 'center' }


      let C33 = worksheetPg.getCell('C33');
      C33.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      C33.alignment = { vertical: 'middle', horizontal: 'left' }


      let C34 = worksheetPg.getCell('C34');
      C34.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C34.alignment = { vertical: 'middle', horizontal: 'left' }


      let L33 = worksheetPg.getCell('L33');
      L33.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      L33.alignment = { vertical: 'middle', horizontal: 'center' }


      let C39 = worksheetPg.getCell('C39');
      C39.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      C39.alignment = { vertical: 'middle', horizontal: 'left' }


      let C44 = worksheetPg.getCell('C44');
      C44.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      C44.alignment = { vertical: 'middle', horizontal: 'left' }


      let J44 = worksheetPg.getCell('J44');
      J44.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      J44.alignment = { vertical: 'middle', horizontal: 'center' }


      let C49 = worksheetPg.getCell('C49');
      C49.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      C49.alignment = { vertical: 'middle', horizontal: 'left' }

      let N49 = worksheetPg.getCell('N49');
      N49.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      N49.alignment = { vertical: 'middle', horizontal: 'center' }


      let C53 = worksheetPg.getCell('C53');
      C53.font = {
        name: 'Cambria',
        size: 12,
        underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      C53.alignment = { vertical: 'middle', horizontal: 'center' }

      let C57 = worksheetPg.getCell('C57');
      C57.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      C57.alignment = { vertical: 'middle', horizontal: 'center' }

      let W57 = worksheetPg.getCell('W57');
      W57.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      W57.alignment = { vertical: 'middle', horizontal: 'center' }

      let C60 = worksheetPg.getCell('C60');
      C60.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C60.alignment = { vertical: 'middle', horizontal: 'left' }


      let C62 = worksheetPg.getCell('C62');
      C62.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C62.alignment = { vertical: 'middle', horizontal: 'left' }


      let C64 = worksheetPg.getCell('C64');
      C64.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C64.alignment = { vertical: 'middle', horizontal: 'left' }


      let C66 = worksheetPg.getCell('C66');
      C66.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C66.alignment = { vertical: 'middle', horizontal: 'left' }

      let C68 = worksheetPg.getCell('C68');
      C68.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C68.alignment = { vertical: 'middle', horizontal: 'left' }

      let C70 = worksheetPg.getCell('C70');
      C70.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C70.alignment = { vertical: 'middle', horizontal: 'left' }


      let C71 = worksheetPg.getCell('C71');
      C71.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C71.alignment = { vertical: 'middle', horizontal: 'left' }


      let S60 = worksheetPg.getCell('S60');
      S60.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      S60.alignment = { vertical: 'middle', horizontal: 'left' }
      S60.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let S62 = worksheetPg.getCell('S62');
      S62.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      S62.alignment = { vertical: 'middle', horizontal: 'left' }
      S62.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let S64 = worksheetPg.getCell('S64');
      S64.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      S64.alignment = { vertical: 'middle', horizontal: 'left' }
      S64.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let S66 = worksheetPg.getCell('S66');
      S66.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      S66.alignment = { vertical: 'middle', horizontal: 'left' }
      S66.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let S68 = worksheetPg.getCell('S68');
      S68.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      S68.alignment = { vertical: 'middle', horizontal: 'left' }
      S68.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let V60 = worksheetPg.getCell('W60');
      V60.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      V60.alignment = { vertical: 'middle', horizontal: 'center' }

      let V64 = worksheetPg.getCell('W64');
      V64.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      V64.alignment = { vertical: 'middle', horizontal: 'center' }

      let V68 = worksheetPg.getCell('W68');
      V68.font = {
        name: 'Cambria',
        size: 12,
       // underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      V68.alignment = { vertical: 'middle', horizontal: 'center' }



      /// fin page de garde


    /// debut fiche de renseignement r1

    const worksheetFr1 = workbook.addWorksheet('Fiche de renseignement R1',{views: [{showGridLines: false}]});

    worksheetFr1.mergeCells('J3', 'AM3');
    worksheetFr1.mergeCells('J6', 'Y6');
    worksheetFr1.mergeCells('DA6', 'AM6');
    worksheetFr1.mergeCells('J9', 'P9');
    worksheetFr1.mergeCells('V9', 'AB9');
    worksheetFr1.mergeCells('AH9', 'AM9');
    worksheetFr1.mergeCells('Y12', 'AD12');
    worksheetFr1.mergeCells('AH12', 'AM12');
    worksheetFr1.mergeCells('U15', 'Z15');
    worksheetFr1.mergeCells('Q19', 'V19');
    worksheetFr1.mergeCells('F23', 'H23');
    worksheetFr1.mergeCells('I23', 'P23');
    worksheetFr1.mergeCells('W23', 'AK23');
    worksheetFr1.mergeCells('F24', 'H24');
    worksheetFr1.mergeCells('I24', 'P24');
    worksheetFr1.mergeCells('W24', 'AK24');
    worksheetFr1.mergeCells('G27', 'N27');
    worksheetFr1.mergeCells('S27', 'Z27');
    worksheetFr1.mergeCells('AE27', 'AL27');
    worksheetFr1.mergeCells('G28', 'N28');
    worksheetFr1.mergeCells('S28', 'Z28');
    worksheetFr1.mergeCells('AE28', 'AL28');
    worksheetFr1.mergeCells('A31', 'AD31');
    worksheetFr1.mergeCells('AE31', 'AM31');
    worksheetFr1.mergeCells('A32', 'AD32');
    worksheetFr1.mergeCells('AE32', 'AM32');
    worksheetFr1.mergeCells('G35', 'N35');
    worksheetFr1.mergeCells('Q35', 'V35');
    worksheetFr1.mergeCells('Y35', 'Z35');
    worksheetFr1.mergeCells('AA35', 'AC35');
    worksheetFr1.mergeCells('AE35', 'AM35');
    worksheetFr1.mergeCells('AA36', 'AC36');
    worksheetFr1.mergeCells('Q36', 'V36');
    worksheetFr1.mergeCells('AE36', 'AM36');
    worksheetFr1.mergeCells('D39', 'AM39');
    worksheetFr1.mergeCells('D40', 'AM40');
    worksheetFr1.mergeCells('D43', 'AD43');
    worksheetFr1.mergeCells('AE43', 'AM43');
    worksheetFr1.mergeCells('D44', 'AD44');
    worksheetFr1.mergeCells('AE44', 'AM44');
    worksheetFr1.mergeCells('D47', 'AM47');
    worksheetFr1.mergeCells('D48', 'AM48');
    worksheetFr1.mergeCells('D51', 'AM51');
    worksheetFr1.mergeCells('D52', 'AM52');
    worksheetFr1.mergeCells('D55', 'AM55');
    worksheetFr1.mergeCells('D53','D54');



    worksheetFr1.columns = [
      {width: 4},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 2},
      {width: 3},
      {width: 3},
      {width: 3},
      {width: 4},
      {width: 3},
      {width: 3},
      {width: 3}];


      let C3 = worksheetFr1.getCell('C3');
      C3.value = "Désignation de l'entreprise";
      C3.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      C3.alignment = { vertical: 'middle', horizontal: 'left' }

      let J3 = worksheetFr1.getCell('J3');
      J3.value = "SEN AUTOMOBILE";
      J3.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      J3.alignment = { vertical: 'middle', horizontal: 'center' }

      let D6 = worksheetFr1.getCell('D6');
      D6.value = "Adresse de l'entreprise";
      D6.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      D6.alignment = { vertical: 'middle', horizontal: 'left' }

      let E9 = worksheetFr1.getCell('E9');
      E9.value = "N° d'identification";
      E9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      E9.alignment = { vertical: 'middle', horizontal: 'left' }


      let J6 = worksheetFr1.getCell('J6');
      J6.value = "145 RUE DE L'ARTISANAT";
      J6.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      J6.alignment = { vertical: 'middle', horizontal: 'center' }


      let AA6 = worksheetFr1.getCell('AA6');
      AA6.value = "Sigle usuel";
      AA6.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      AA6.alignment = { vertical: 'middle', horizontal: 'left' }

      let DA6 = worksheetFr1.getCell('DA6');
      DA6.value = "0";
      DA6.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      DA6.alignment = { vertical: 'middle', horizontal: 'center' }


      let R9 = worksheetFr1.getCell('R9');
      R9.value = "Exercice clos le";
      R9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      R9.alignment = { vertical: 'middle', horizontal: 'left' }


      let V9 = worksheetFr1.getCell('V9');
      V9.value = "31/12/2019";
      V9.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      V9.alignment = { vertical: 'middle', horizontal: 'center' }


      let AC9 = worksheetFr1.getCell('AC9');
      AC9.value = "Durée (en mois)";
      AC9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      AC9.alignment = { vertical: 'middle', horizontal: 'left' }


      let AH9 = worksheetFr1.getCell('AH9');
      AH9.value = "12";
      AH9.font = {
        name: 'Cambria',
        size: 12,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      AH9.alignment = { vertical: 'middle', horizontal: 'center' }


      let B12 = worksheetFr1.getCell('B12');
      B12.value = "ZA";
      B12.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B12.alignment = { vertical: 'middle', horizontal: 'center' }
      B12.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let B15 = worksheetFr1.getCell('B15');
      B15.value = "ZB";
      B15.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }
      B15.alignment = { vertical: 'middle', horizontal: 'center' }
      B15.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let B19 = worksheetFr1.getCell('B19');
      B19.value = "ZC";
      B19.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B19.alignment = { vertical: 'middle', horizontal: 'center' }
      B19.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let B23 = worksheetFr1.getCell('B23');
      B23.value = "ZD";
      B23.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B23.alignment = { vertical: 'middle', horizontal: 'center' }
      B23.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let B27 = worksheetFr1.getCell('B27');
      B27.value = "ZE";
      B27.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }
      B27.alignment = { vertical: 'middle', horizontal: 'center' }
      B27.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let B31 = worksheetFr1.getCell('B31');
      B31.value = "ZF";
      B31.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }
      B31.alignment = { vertical: 'middle', horizontal: 'center' }
      B31.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }



      let B35 = worksheetFr1.getCell('B35');
      B35.value = "ZG";
      B35.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B35.alignment = { vertical: 'middle', horizontal: 'center' }
      B35.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let B39 = worksheetFr1.getCell('B39');
      B39.value = "ZH";
      B39.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B39.alignment = { vertical: 'middle', horizontal: 'center' }
      B39.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }



      let B43 = worksheetFr1.getCell('B43');
      B43.value = "ZI";
      B43.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B43.alignment = { vertical: 'middle', horizontal: 'center' }
      B43.border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let E12 = worksheetFr1.getCell('E12');
      E12.value = "EXERCICE COMPTABLE";
      E12.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      E12.alignment = { vertical: 'middle', horizontal: 'left' }


      let W12 = worksheetFr1.getCell('W12');
      W12.value = "DU";
      W12.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      W12.alignment = { vertical: 'middle', horizontal: 'left' }


      let Y12 = worksheetFr1.getCell('Y12');
      Y12.value = "01/01/2019";
      Y12.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      Y12.alignment = { vertical: 'middle', horizontal: 'center' }


      let AF12 = worksheetFr1.getCell('AF12');
     AF12.value = "AU";
     AF12.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
     AF12.alignment = { vertical: 'middle', horizontal: 'left' }


     let AH12 = worksheetFr1.getCell('AH12');
      AH12.value = "01/01/2019";
      AH12.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      AH12.alignment = { vertical: 'middle', horizontal: 'center' }


      let E15 = worksheetFr1.getCell('E15');
      E15.value = "DATE D'ARRETE EFFECTIF DES COMPTES";
      E15.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      E15.alignment = { vertical: 'middle', horizontal: 'left' }


      let U15 = worksheetFr1.getCell('U15');
      U15.value = "01/01/2019";
      U15.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      U15.alignment = { vertical: 'middle', horizontal: 'center' }


      let E19 = worksheetFr1.getCell('E19');
      E19.value = "EXERCICE PRECEDENT CLOS LE";
      E19.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      E19.alignment = { vertical: 'middle', horizontal: 'left' }


      let Q19 = worksheetFr1.getCell('Q19');
      Q19.value = "aucun";
      Q19.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      Q19.alignment = { vertical: 'middle', horizontal: 'center' }



      let AA19 = worksheetFr1.getCell('AA19');
      AA19.value = "DUREE EXERCICE PRECEDENT EN MOIS :";
      AA19.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      AA19.alignment = { vertical: 'middle', horizontal: 'left' }
     
      let AL19 = worksheetFr1.getCell('AL19');
      AL19.value = "12";
      AL19.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      AL19.alignment = { vertical: 'middle', horizontal: 'center' }


      let F23 = worksheetFr1.getCell('F23');
      F23.value = "SN";
      F23.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      F23.alignment = { vertical: 'middle', horizontal: 'center' }

      let I23 = worksheetFr1.getCell('I23');
      I23.value = "DKR.2018.B.2002";
      I23.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      I23.alignment = { vertical: 'middle', horizontal: 'center' }

      let F24 = worksheetFr1.getCell('F24');
      F24.value = "Greffe";
      F24.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      F24.alignment = { vertical: 'middle', horizontal: 'center' }

      let I24 = worksheetFr1.getCell('I24');
      I24.value = "N° registre du Commerce";
      I24.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      I24.alignment = { vertical: 'middle', horizontal: 'center' }


      let W24 = worksheetFr1.getCell('W24');
      W24.value = "N° répertoire des entreprises";
      W24.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      W24.alignment = { vertical: 'middle', horizontal: 'center' }

      let G27 = worksheetFr1.getCell('G27');
      G27.value = "55552212121";
      G27.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      G27.alignment = { vertical: 'middle', horizontal: 'center' }

      let G28 = worksheetFr1.getCell('G28');
      G28.value = "N° de sécurité sociale";
      G28.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      G28.alignment = { vertical: 'middle', horizontal: 'center' }

      let S28 = worksheetFr1.getCell('S28');
      S28.value = "N° Code employeur";
      S28.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      S28.alignment = { vertical: 'middle', horizontal: 'center' }

      let AE28 = worksheetFr1.getCell('AE28');
      AE28.value = "Code activité principale";
      AE28.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      AE28.alignment = { vertical: 'middle', horizontal: 'center' }


      let D31 = worksheetFr1.getCell('D31');
      D31.value = "SEN AUTOMOBILE";
      D31.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      D31.alignment = { vertical: 'middle', horizontal: 'center' }

      let AE31 = worksheetFr1.getCell('AE31');
      AE31.value = "0";
      AE31.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      AE31.alignment = { vertical: 'middle', horizontal: 'center' }

      let A32 = worksheetFr1.getCell('A32');
      A32.value = "Désignation de l'entreprise";
      A32.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      A32.alignment = { vertical: 'middle', horizontal: 'center' }

      let AE32 = worksheetFr1.getCell('AE32');
      AE32.value = "Sigle";
      AE32.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      AE32.alignment = { vertical: 'middle', horizontal: 'center' }

      let G35 = worksheetFr1.getCell('G35');
      G35.value = "77 639 34 38";
      G35.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      G35.alignment = { vertical: 'middle', horizontal: 'center' }

      let AE35 = worksheetFr1.getCell('AE35');
      AE35.value = "Dakar";
      AE35.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      AE35.alignment = { vertical: 'middle', horizontal: 'center' }




      let G36 = worksheetFr1.getCell('G36');
      G36.value = "N° de téléphone";
      G36.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }
      G36.alignment = { vertical: 'middle', horizontal: 'center' }

      let Q36 = worksheetFr1.getCell('Q36');
      Q36.value = "Télécopie";
      Q36.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }
      Q36.alignment = { vertical: 'middle', horizontal: 'center' }


      let Y36 = worksheetFr1.getCell('Y36');
      Y36.value = "Code";
      Y36.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }
      Y36.alignment = { vertical: 'middle', horizontal: 'center' }

      let AA36 = worksheetFr1.getCell('AA36');
      AA36.value = "Boîte postale";
      AA36.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }
      AA36.alignment = { vertical: 'middle', horizontal: 'center' }

    
      let AE36 = worksheetFr1.getCell('AE36');
      AE36.value = "Ville";
      AE36.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      AE36.alignment = { vertical: 'middle', horizontal: 'center' }


      let D39 = worksheetFr1.getCell('D39');
      D39.value = "145 RUE DE L'ARTISANAT";
      D39.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      D39.alignment = { vertical: 'middle', horizontal: 'center' }

      let D40 = worksheetFr1.getCell('D40');
      D40.value = "Adresse géographique complète (Immeuble, rue, quartier, ville, pays)";
      D40.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      D40.alignment = { vertical: 'middle', horizontal: 'center' }


      let D43 = worksheetFr1.getCell('D43');
      D43.value = "AUDIT ET CONSEIL  544 54545";
      D43.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      D43.alignment = { vertical: 'middle', horizontal: 'center' }


      let D44 = worksheetFr1.getCell('D44');
      D44.value = "Désignation précise de l'activité principale exercée par l'entreprise";
      D44.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      D44.alignment = { vertical: 'middle', horizontal: 'center' }


      let AE44 = worksheetFr1.getCell('AE44');
      AE44.value = "% capacité production utilisée";
      AE44.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      AE44.alignment = { vertical: 'middle', horizontal: 'center' }


      let D47 = worksheetFr1.getCell('D47');
      D47.value = "MBA";
      D47.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      D47.alignment = { vertical: 'middle', horizontal: 'center' }

      let D48 = worksheetFr1.getCell('D48');
      D48.value = "Nom, adresse et qualité de personne à contacter en cas de demande d'informations complémentaires";
      D48.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      D48.alignment = { vertical: 'middle', horizontal: 'center' }

      let D51 = worksheetFr1.getCell('D51');
      D51.value = "DFA";
      D51.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      D51.alignment = { vertical: 'middle', horizontal: 'center' }

      let D52 = worksheetFr1.getCell('D52');
      D52.value = "Nom du professionnel salarié de l'entreprise ou Nom, adresse et téléphone du cabinet comptable ou du professionnel INSCRIT A L'ORDRE NATIONAL DES EXPERTS COMPTABLES ET DES COMPTABLES AGREES ayant établi les états financiers.";
      D52.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      D52.alignment = { vertical: 'middle', horizontal: 'center' }



      let D55 = worksheetFr1.getCell('D55');
      D55.value = "DAO";
      D55.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      D55.alignment = { vertical: 'middle', horizontal: 'center' }

      let D56 = worksheetFr1.getCell('D56');
      D56.value = "Noms et adresses du ou des commissaires au comptes";
      D56.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      D56.alignment = { vertical: 'middle', horizontal: 'center' }


      /// fin fiche de renseignement 


      /// debut Activités de l'entreprise R2

      const worksheetFr2 = workbook.addWorksheet('Activités de l\'entreprise R2',{views: [{showGridLines: false}]});
      

    worksheetFr2.mergeCells('C1','S1');
    worksheetFr2.mergeCells('B3','K3');
    worksheetFr2.mergeCells('M3','S3');
    worksheetFr2.mergeCells('B5','I5');
    worksheetFr2.mergeCells('N5','S5');
    worksheetFr2.mergeCells('L7','Q7');
    worksheetFr2.mergeCells('A23','T23');
    worksheetFr2.mergeCells('C25','J25');
    worksheetFr2.mergeCells('K25','L25');


    const R2Data = [
      {
        "lignes": 45
      },
      {
"B1":"Désignation de l'entreprise :",
"C1":"ENTREPRISE MORAL",														
"B2":"Adresse de l'entreprise :",															
"B3":"ZONE 15",										
"L3":"Sigle usuel :",	
"M3":"0",				
"B4":"N° d'identification :",										
"L4":"Exercice clos le :",					
"B5":"04942901",										
"L5":"31/12/2019",	
"M5":"Durée (en mois)",
"N5":"12",			
"L7":"Contrôle de l'entreprise (cocher la case)",				
"A9":"ZK",	
"B9":"Forme juridique (1)",				
"F9":"0",	
"G9":"2",			
"K9":"ZQ",	
"L9":"Entreprise sous contrôle public",					
"A11":"ZL",	
"B11":"Régime fiscal  (1)",				
"F11":"2",				
"K11":"ZR",	
"L11":"Entreprise sous contrôle privé national",					
"Q11":"X",
"A13":"ZM",	
"B13":"Pays du siège social (1)",			
"F13":"0",
"G13":"7",			
"K13":"ZS",	
"L13":"Entreprise sous contrôle privé étranger",					
"A15":"ZN",	
"B15":"Nbre d'établissements dans le pays",				
"F15":"0",	
"G15":"1",									
"A17":"ZO",	
"B17":"Nbre d'établissements hors du pays",				
"F17":"0",
"G17":"0",									
"B18":"pour lesquels une comptabilité distincte est tenue",															
"A20":"ZP",	
"B20":"1ère année exercice dans le pays",		
"F20":"2",
"G20":"0",
"H20":"0",
"I20":"1",
"A23":"ACTIVITE DE L'ENTREPRISE",															
"B25":"DESIGNATION DE L'ACTIVITE (2)",	
"C25":"Code nomenclature \n d'activité (1)",								
"K25":"Valeur Ajoutée (VA) HT",
"M25":"% activité dans le \n CA HT ou la VA",				
"B28":"COMMERCE",
"C28":"0",
"D28":"0",
"E28":"3",
"F28":"1",
"G28":"0",
"H28":"0",
"I28":"3",			
"M28":"100 %",			
"M31":"#DIV/0 !",				
"M34":"#DIV/0 !",				
"M37":"#DIV/0 !",				
"M40":"#DIV/0 !",			
"M43":"#DIV/0 !",				
"M46":"#DIV/0 !",				
"M49":"#DIV/0 !",				
"B52":"Divers",															
"H54":"TOTAL",		 
"L54":"-",   	
"M54":"#DIV/0 !",			
"B57":"(1) Se réréfer aux tables des codes",															
"B58":"(2) Lister de manière précise les activités dans l'ordre décroissant du CAHT, ou de la valeur ajoutée (V.A.)",														
"B59":"(3) Rayer la mention inutile (utiliser de préférence la V.A.)"
      }
    ]



  const  alphR2: string[] = [
      'A',
      'B',
      'C',
      'D',
      'E',
      'F',
      'G',
      'H',
      'I',
      'J',
      'K',
      'L',
      'M',
      'N',
      'O',
      'P',
      'Q',
      'R',
      'S'];

      const  alphR2Bord: string[] = [
        'C',
        'D',
        'E',
        'F',
        'G',
        'H',
        'I',
        'J',
        'K',
        'L',
        'M',
        'N',
        'O',
        'P',
        'Q',
        'R',
        'S'];


        const  alphR2Bord2: string[] = [
          'K',
          'L',
          'M',
          'N',
          'O',
          'P',
          'Q',
          'R',
          'S'];

        


        alphR2Bord.forEach((item) => {
          worksheetFr2.getCell(`${item}1`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            //left: {style:'hair', color: {argb:'000000'}},
            //right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
  
        })

        alphR2Bord2.forEach((item) => {
          worksheetFr2.getCell(`${item}7`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            //left: {style:'hair', color: {argb:'000000'}},
            //right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
  
        })


          worksheetFr2.getCell(`M3`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            //left: {style:'hair', color: {argb:'000000'}},
            //right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
  
        


       
          worksheetFr2.getCell(`B3`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            //left: {style:'hair', color: {argb:'000000'}},
            //right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }


          worksheetFr2.getCell(`F9`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
          worksheetFr2.getCell(`G9`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }

          worksheetFr2.getCell(`F11`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }


          worksheetFr2.getCell(`F13`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }

          worksheetFr2.getCell(`G13`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }

          worksheetFr2.getCell(`F15`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }

          worksheetFr2.getCell(`G15`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }


          worksheetFr2.getCell(`F17`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }

          worksheetFr2.getCell(`G17`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }


          worksheetFr2.getCell(`F20`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }

          worksheetFr2.getCell(`G20`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
          worksheetFr2.getCell(`H20`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }

          worksheetFr2.getCell(`I20`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }







  

          worksheetFr2.getCell(`B5`).border = {
            bottom: {style:'hair', color: {argb:'000000'}},
            //left: {style:'hair', color: {argb:'000000'}},
            //right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
  
      


      

      alphR2.forEach((item) => {
        worksheetFr2.getCell(`${item}6`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
          //left: {style:'hair', color: {argb:'000000'}},
          //right: {style:'hair', color: {argb:'000000'}},
          //top: {style:'hair', color: {argb:'000000'}}
        }

      })


      alphR2.forEach((item) => {
        worksheetFr2.getCell(`${item}21`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
          //left: {style:'hair', color: {argb:'000000'}},
          //right: {style:'hair', color: {argb:'000000'}},
          //top: {style:'hair', color: {argb:'000000'}}
        }

      })

      alphR2.forEach((item) => {
        worksheetFr2.getCell(`${item}6`).border = {
          bottom: {style:'hair', color:  {argb:'000000'}},
          //left: {style:'hair', color: {argb:'000000'}},
          //right: {style:'hair', color: {argb:'000000'}},
          //top: {style:'hair', color: {argb:'000000'}}
        }

      })

     



       






      worksheetFr2.columns = [
        {width: 4},
        {width: 50},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 22},
        {width: 18},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2},
        {width: 2}];

       worksheetFr2.getRow(25).height = 40;



       for(let i = 0; i <= Object.keys(R2Data).length; i++ ){

        for(var v in R2Data[1]){
    
        
             console.log(`mon ${v}`)
              let cell = worksheetFr2.getCell(v);
             cell.value = R2Data[1][v];
              cell.font = {
                name: 'Cambria',
                size: 9,
                //underline: 'single',
                bold: v == "H54"? true : false,
                color: { argb: '000000' }
              }
        
              cell.alignment = { vertical: 'middle', horizontal: 'left'}
    
            
            
      
        }
    
      }


      let B1 = worksheetFr2.getCell('B1');
      B1.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B1.alignment = { vertical: 'middle', horizontal: 'left' }


      let C1 = worksheetFr2.getCell('C1');
      C1.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      C1.alignment = { vertical: 'middle', horizontal: 'center' }

      let B2 = worksheetFr2.getCell('B2');
      B2.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B2.alignment = { vertical: 'middle', horizontal: 'center' }


      let BR23 = worksheetFr2.getCell('B3');
      BR23.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      BR23.alignment = { vertical: 'middle', horizontal: 'center' }


      let L3 = worksheetFr2.getCell('L3');
      L3.value = "Sigle usuel :";
      L3.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      L3.alignment = { vertical: 'middle', horizontal: 'center' }

      let M3 = worksheetFr2.getCell('M3');
      M3.value = "0";
      M3.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      M3.alignment = { vertical: 'middle', horizontal: 'center' }

      let B4 = worksheetFr2.getCell('B4');
      B4.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B4.alignment = { vertical: 'middle', horizontal: 'center' }



      let L4 = worksheetFr2.getCell('L4');
      L4.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      L4.alignment = { vertical: 'middle', horizontal: 'center' }

      let BR25 = worksheetFr2.getCell('B5');
      BR25.value = "12";
      BR25.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      BR25.alignment = { vertical: 'middle', horizontal: 'center' }


      let L5 = worksheetFr2.getCell('L5');
      L5.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      L5.alignment = { vertical: 'middle', horizontal: 'center' }


      let M5 = worksheetFr2.getCell('M5');
      M5.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      M5.alignment = { vertical: 'middle', horizontal: 'center' }



      let A9 = worksheetFr2.getCell('A9');
      A9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      A9.alignment = { vertical: 'middle', horizontal: 'center' }


      let A11 = worksheetFr2.getCell('A11');
      A11.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      A11.alignment = { vertical: 'middle', horizontal: 'center' }



      let A13 = worksheetFr2.getCell('A13');
      A13.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      A13.alignment = { vertical: 'middle', horizontal: 'center' }


      let AR215 = worksheetFr2.getCell('A15');
      AR215.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      AR215.alignment = { vertical: 'middle', horizontal: 'center' }


      let A17 = worksheetFr2.getCell('A17');
      A17.value = "ZO";
      A17.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      A17.alignment = { vertical: 'middle', horizontal: 'center' }


      let A20 = worksheetFr2.getCell('A20');
      A20.value = "ZP";
      A20.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      A20.alignment = { vertical: 'middle', horizontal: 'center' }



      let B9 = worksheetFr2.getCell('B9');
      B9.value = "Forme juridique (1)";
      B9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B9.alignment = { vertical: 'middle', horizontal: 'left' }


      let B11 = worksheetFr2.getCell('B11');
      B11.value = "Régime fiscal  (1)";
      B11.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B11.alignment = { vertical: 'middle', horizontal: 'left' }



      let B13 = worksheetFr2.getCell('B13');
      B13.value = "Pays du siège social (1)";
      B13.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B13.alignment = { vertical: 'middle', horizontal: 'left' }


      let BR215 = worksheetFr2.getCell('B15');
      BR215.value = "Nbre d'établissements dans le pays";
      BR215.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      BR215.alignment = { vertical: 'middle', horizontal: 'left' }


      let B17 = worksheetFr2.getCell('B17');
      B17.value = "Nbre d'établissements hors du pays";
      B17.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B17.alignment = { vertical: 'middle', horizontal: 'left' }

      let B18 = worksheetFr2.getCell('B18');
      B18.value = "pour lesquels une comptabilité distincte est tenue";
      B18.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B18.alignment = { vertical: 'middle', horizontal: 'left' }


      let B20 = worksheetFr2.getCell('B20');
      B20.value = "1ère année exercice dans le pays";
      B20.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      B20.alignment = { vertical: 'middle', horizontal: 'left' }


      let F9 = worksheetFr2.getCell('F9');
      F9.value = "0";
      F9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      F9.alignment = { vertical: 'middle', horizontal: 'center' }


      let G9 = worksheetFr2.getCell('G9');
      G9.value = "1";
      G9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      G9.alignment = { vertical: 'middle', horizontal: 'center' }

      let F11 = worksheetFr2.getCell('F11');
      F11.value = "0";
      F11.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      F11.alignment = { vertical: 'middle', horizontal: 'center' }


      let F13 = worksheetFr2.getCell('F13');
      F13.value = "0";
      F13.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      F13.alignment = { vertical: 'middle', horizontal: 'center' }


      let G13 = worksheetFr2.getCell('G13');
      G13.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      G13.alignment = { vertical: 'middle', horizontal: 'center' }


      let F15 = worksheetFr2.getCell('F15');
      F15.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      F15.alignment = { vertical: 'middle', horizontal: 'center' }


      let G15 = worksheetFr2.getCell('G15');
      G15.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      G15.alignment = { vertical: 'middle', horizontal: 'center' }

      let F17 = worksheetFr2.getCell('F17');
      F17.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      F17.alignment = { vertical: 'middle', horizontal: 'center' }


      let G17 = worksheetFr2.getCell('G17');
      G17.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      G17.alignment = { vertical: 'middle', horizontal: 'center' }


      let E20 = worksheetFr2.getCell('E20');
      E20.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      E20.alignment = { vertical: 'middle', horizontal: 'center' }

      let F20 = worksheetFr2.getCell('F20');
      F20.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      F20.alignment = { vertical: 'middle', horizontal: 'center' }

      let G20 = worksheetFr2.getCell('G20');
      G20.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      G20.alignment = { vertical: 'middle', horizontal: 'center' }


      let H20 = worksheetFr2.getCell('H20');
      H20.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      H20.alignment = { vertical: 'middle', horizontal: 'center' }


      let L7 = worksheetFr2.getCell('L7');
      L7.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      L7.alignment = { vertical: 'middle', horizontal: 'center' }


      let K9 = worksheetFr2.getCell('K9');
      K9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      K9.alignment = { vertical: 'middle', horizontal: 'center' }


      let K11 = worksheetFr2.getCell('K11');
      K11.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      K11.alignment = { vertical: 'middle', horizontal: 'center' }

      let K13 = worksheetFr2.getCell('K13');
      K13.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      K13.alignment = { vertical: 'middle', horizontal: 'center' }



      let L9 = worksheetFr2.getCell('L9');
      L9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      L9.alignment = { vertical: 'middle', horizontal: 'left' }


      let L11 = worksheetFr2.getCell('L11');
      L11.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      L11.alignment = { vertical: 'middle', horizontal: 'center' }

      let L13 = worksheetFr2.getCell('L13');
      L13.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      L13.alignment = { vertical: 'middle', horizontal: 'center' }

      let Q9 = worksheetFr2.getCell('Q9');
      Q9.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      Q9.alignment = { vertical: 'middle', horizontal: 'center' }

      Q9.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let Q11 = worksheetFr2.getCell('Q11');
      Q11.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      Q11.alignment = { vertical: 'middle', horizontal: 'center' }

      Q11.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let Q13 = worksheetFr2.getCell('Q13');
      Q13.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      Q13.alignment = { vertical: 'middle', horizontal: 'center' }

      Q13.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let A23 = worksheetFr2.getCell('A23');
      A23.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      A23.alignment = { vertical: 'middle', horizontal: 'center' }


      let BR225 = worksheetFr2.getCell('B25');
      BR225.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      BR225.alignment = { vertical: 'middle', horizontal: 'center' }
      BR225.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      let C25 = worksheetFr2.getCell('C25');
      C25.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      C25.alignment = { vertical: 'middle', horizontal: 'justify' }
      C25.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let K25 = worksheetFr2.getCell('K25');
      K25.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      K25.alignment = { vertical: 'middle', horizontal: 'justify' }
      K25.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }


      let MR225 = worksheetFr2.getCell('M25');
      MR225.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      MR225.alignment = { vertical: 'middle', horizontal: 'center' }

      MR225.border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      for(let i = 26; i < 56; i++){
        worksheetFr2.getCell(`B${i}`).border = {
          left: {style:'hair', color: {argb:'000000'}},
          right: {style:'hair', color: {argb:'000000'}},
          
        }
      }

      for(let i = 26; i < 56; i++){
        worksheetFr2.getCell(`J${i}`).border = {
          right: {style:'hair', color: {argb:'000000'}},
          
        }
      }

      for(let i = 26; i < 56; i++){
        worksheetFr2.getCell(`L${i}`).border = {
          right: {style:'hair', color: {argb:'000000'}},
          
        }
      }

      for(let i = 26; i < 56; i++){
        worksheetFr2.getCell(`M${i}`).border = {
          right: {style:'hair', color: {argb:'000000'}},
          
        }
      }

      worksheetFr2.getCell(`M55`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},

        
      }
      worksheetFr2.getCell(`B55`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},

        
        
      }

      worksheetFr2.getCell(`K55`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        
      }
      worksheetFr2.getCell(`L55`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},


        
      }

      const alphBord = [
        'B',
        'C',
        'D',
        'E',
        'F',
        'G',
        'H',
        'I',
        'J',
        'K',
        'L',
        'M'];
      const alphBord2 = [
          'D',
          'E',
          'F',
          'G',
          'H',
          'I'];

          alphBord2.forEach((item) => {
            worksheetFr2.getCell(`${item}28`).border = {
              bottom: {style:'hair', color:  {argb:'000000'}},
              left: {style:'hair', color:  {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              //top: {style:'hair', color: {argb:'000000'}}
            }
    
          })

          alphBord2.forEach((item) => {
            worksheetFr2.getCell(`${item}31`).border = {
              bottom: {style:'hair', color:  {argb:'000000'}},
              left: {style:'hair', color:  {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              //top: {style:'hair', color: {argb:'000000'}}
            }
    
          })

          alphBord2.forEach((item) => {
            worksheetFr2.getCell(`${item}34`).border = {
              bottom: {style:'hair', color:  {argb:'000000'}},
              left: {style:'hair', color:  {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              //top: {style:'hair', color: {argb:'000000'}}
            }
    
          })

          alphBord2.forEach((item) => {
            worksheetFr2.getCell(`${item}37`).border = {
              bottom: {style:'hair', color:  {argb:'000000'}},
              left: {style:'hair', color:  {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              //top: {style:'hair', color: {argb:'000000'}}
            }
    
          })

          alphBord2.forEach((item) => {
            worksheetFr2.getCell(`${item}40`).border = {
              bottom: {style:'hair', color:  {argb:'000000'}},
              left: {style:'hair', color:  {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              //top: {style:'hair', color: {argb:'000000'}}
            }
    
          })

          alphBord2.forEach((item) => {
            worksheetFr2.getCell(`${item}43`).border = {
              bottom: {style:'hair', color:  {argb:'000000'}},
              left: {style:'hair', color:  {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              //top: {style:'hair', color: {argb:'000000'}}
            }
    
          })

          alphBord2.forEach((item) => {
            worksheetFr2.getCell(`${item}46`).border = {
              bottom: {style:'hair', color:  {argb:'000000'}},
              left: {style:'hair', color:  {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              //top: {style:'hair', color: {argb:'000000'}}
            }
    
          })

          alphBord2.forEach((item) => {
            worksheetFr2.getCell(`${item}49`).border = {
              bottom: {style:'hair', color:  {argb:'000000'}},
              left: {style:'hair', color:  {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              //top: {style:'hair', color: {argb:'000000'}}
            }
    
          })


        alphBord.forEach((item) => {
          worksheetFr2.getCell(`${item}51`).border = {
            bottom: {style:'hair', color:  {argb:'000000'}},
            left: {style:'hair', color:  {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
  
        })

        alphBord.forEach((item) => {
          worksheetFr2.getCell(`${item}53`).border = {
            bottom: {style:'hair', color:  {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
  
        })

        alphBord.forEach((item) => {
          worksheetFr2.getCell(`${item}55`).border = {
            bottom: {style:'hair', color:  {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            //top: {style:'hair', color: {argb:'000000'}}
          }
  
        })

         for(let i = 7; i <= 21; i++){ worksheetFr2.getCell(`S${i}`).border = {
          //bottom: {style:'hair', color:  {argb:'000000'}},
          //left: {style:'hair', color: {argb:'000000'}},
          right: {style:'hair', color: {argb:'000000'}},
          //top: {style:'hair', color: {argb:'000000'}}
        }



         }

         for(let i = 7; i <= 21; i++){ worksheetFr2.getCell(`J${i}`).border = {
          //bottom: {style:'hair', color:  {argb:'000000'}},
          //left: {style:'hair', color: {argb:'000000'}},
          right: {style:'hair', color: {argb:'000000'}},
          //top: {style:'hair', color: {argb:'000000'}}
        }



         }
       

        

  



      /// fin  Activités de l'entreprise R2




      /// debut Dirigeants R3

      const worksheetFr3 = workbook.addWorksheet('Dirigeants R3',{views: [{showGridLines: false}]});

    worksheetFr3.mergeCells('B1','F1');
    worksheetFr3.mergeCells('B2','C2');
    worksheetFr3.mergeCells('A5','F5');
    worksheetFr3.mergeCells('F7','E7');
    worksheetFr3.mergeCells('A25','F25');
    worksheetFr3.mergeCells('C27','D27');
    worksheetFr3.mergeCells('E27','F27');





    const R3Data =[
      {
        "lignes": 25
      },
      {
"A1":"Désignation de l'entreprise",	
"B1":"ENTREPRISE MORAL",				
"A2":"Adresse de l'entreprise",	
"B2":"ZONE 15",		
"D2":"Sigle usuel",	
"E2":"0",	
"A3":"N° d'identification fiscale",	
"B3":"004942901",	
"C3":"Exercice clos le",	
"D3":"31/12/2019",	
"E3":"Durée (en mois)",	
"F3":"12",
"A5":"DIRIGEANTS (1)",					
"A7":"Nom",	
"B7":"Prénoms",	
"C7":"Qualité",	
"D7":"N° d'identification \n fiscale",	
"E7":"Adressse (BP, ville, pays)",	
"A8":"KASSE",	
"B8":"KHALIDOU", 	
"C8":"GERANT",	
"D8":"1920078",	
"E8":"PARCELLE",	
"A23":"(1) Dirigeants = Président Directeur Général, Directeur Général, Adminsitrateur Général, Gérant, Autres.",					
"A25":"MEMBRES DU CONSEIL D'ADMINISTRATION",					
"A27":"Nom",	
"B27":"Prénoms",	
"C27":"Qualité",		
"E27":"Adressse (BP, ville, pays)",
      }
    ]


    
    for(let i = 0; i <= Object.keys(R3Data).length; i++ ){

      for(var v in R3Data[1]){
  
      
           console.log(`mon ${v}`)
            let cell = worksheetFr3.getCell(v);
           cell.value = R3Data[1][v];
            cell.font = {
              name: 'Cambria',
              size: 9,
              //underline: 'single',
              bold: v == "H54"? true : false,
              color: { argb: '000000' }
            }
      
            cell.alignment = { vertical: 'middle', horizontal: 'left'}
  
          
          
    
      }
  
    }

    
    


    
    worksheetFr3.getRow(7).height = 30


      worksheetFr3.columns = [
        {width: 26},
        {width: 19},
        {width: 19},
        {width: 18},
        {width: 18},
        {width: 13}];

      let A1R3 = worksheetFr3.getCell('A1');
      A1R3.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      A1R3.alignment = { vertical: 'middle', horizontal: 'left' }

      let B1R3 = worksheetFr3.getCell('B1');
      B1R3.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      B1R3.alignment = { vertical: 'middle', horizontal: 'center' }

      let A2R3 = worksheetFr3.getCell('A2');
      A2R3.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        //bold: true,
        color: { argb: '000000' }
      }
      A2R3.alignment = { vertical: 'middle', horizontal: 'left' }

      let B2R3 = worksheetFr3.getCell('B2');
      B2R3.font = {
        name: 'Cambria',
        size: 10,
        //underline: 'single',
        bold: true,
        color: { argb: '000000' }
      }
      B2R3.alignment = { vertical: 'middle', horizontal: 'center' }

      let E2R3 = worksheetFr3.getCell('E2');
       E2R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         bold: true,
         color: { argb: '000000' }
       }
       E2R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let A3R3 = worksheetFr3.getCell('A3');
       A3R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       A3R3.alignment = { vertical: 'middle', horizontal: 'left' }

       let B3R3 = worksheetFr3.getCell('B3');
       B3R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         bold: true,
         color: { argb: '000000' }
       }
       B3R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let C3R3 = worksheetFr3.getCell('C3');
       C3R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       C3R3.alignment = { vertical: 'middle', horizontal: 'center' }


      let D3R3 = worksheetFr3.getCell('D3');
       D3R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         bold: true,
         color: { argb: '000000' }
       }
       D3R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let E3R3 = worksheetFr3.getCell('E3');
       E3R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       E3R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let F3R3 = worksheetFr3.getCell('F3');
       F3R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         bold: true,
         color: { argb: '000000' }
       }
       F3R3.alignment = { vertical: 'middle', horizontal: 'center' }


       let A5R3 = worksheetFr3.getCell('A5');
       A5R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         bold: true,
         color: { argb: '000000' }
       }
       A5R3.alignment = { vertical: 'middle', horizontal: 'center' }


       let A7R3 = worksheetFr3.getCell('A7');
       A7R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       A7R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let B7R3 = worksheetFr3.getCell('B7');
       B7R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       B7R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let C7R3 = worksheetFr3.getCell('C7');
       C7R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       C7R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let D7R3 = worksheetFr3.getCell('D7');
       D7R3.font = {
         name: 'Cambria',
         size: 9,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       D7R3.alignment = { vertical: 'middle', horizontal: 'center' }

       
       let E7R3 = worksheetFr3.getCell('E7');
       E7R3.font = {
         name: 'Cambria',
         size: 9,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       E7R3.alignment = { vertical: 'middle', horizontal: 'center' }


      let alp3 = ['A','B', 'C','D','E','F'];

      alp3.forEach((item) => {
        worksheetFr3.getCell(`${item}7`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
          left: item != "F" ? {style:'hair', color: {argb:'000000'}} : {style:'hair', color: {argb:'FFFFFF'}},
          right: {style:'hair', color: {argb:'000000'}},
          top: {style:'hair', color: {argb:'000000'}}
        }
      })


      for(let i = 8; i <= 22; i++){
       worksheetFr3.getCell(`B${i}`).border =  {
          bottom: i == 22 ? {style:'hair', color: {argb:'000000'}} :{style:'hair', color: {argb:'FFFFFF'}}  ,
          left: {style:'hair', color: {argb:'000000'}},
          right: {style:'hair', color: {argb:'000000'}},
          //top: {style:'hair', color: {argb:'000000'}}
        }
        
      }

      for(let i = 8; i <= 22; i++){
        worksheetFr3.getCell(`C${i}`).border =  {
          bottom: i == 22 ? {style:'hair', color: {argb:'000000'}} :{style:'hair', color: {argb:'FFFFFF'}}  ,
           left: {style:'hair', color: {argb:'000000'}},
           right: {style:'hair', color: {argb:'000000'}},
           //top: {style:'hair', color: {argb:'000000'}}
         }
         
       }


       for(let i = 8; i <= 22; i++){
        worksheetFr3.getCell(`D${i}`).border =  {
          bottom: i == 22 ? {style:'hair', color: {argb:'000000'}} :{style:'hair', color: {argb:'FFFFFF'}}  ,
           left: {style:'hair', color: {argb:'000000'}},
           right: {style:'hair', color: {argb:'000000'}},
           //top: {style:'hair', color: {argb:'000000'}}
         }
         
       }

       for(let i = 8; i <= 22; i++){
        worksheetFr3.getCell(`F${i}`).border =  {
           bottom: i == 22 ? {style:'hair', color: {argb:'000000'}} :{style:'hair', color: {argb:'FFFFFF'}}  ,
           //left: {style:'hair', color: {argb:'000000'}},
           right: {style:'hair', color: {argb:'000000'}},
           //top: {style:'hair', color: {argb:'000000'}}
         }
         
       }

       worksheetFr3.getCell('A22').border =  {
        bottom:  {style:'hair', color: {argb:'000000'}},
         left: {style:'hair', color: {argb:'000000'}},
         right: {style:'hair', color: {argb:'000000'}},
         //top: {style:'hair', color: {argb:'000000'}}
       }

       worksheetFr3.getCell('E22').border =  {
        bottom:  {style:'hair', color: {argb:'000000'}},
         left: {style:'hair', color: {argb:'000000'}},
         //right: {style:'hair', color: {argb:'000000'}},
         //top: {style:'hair', color: {argb:'000000'}}
       }

       let A27R3 = worksheetFr3.getCell('A27');
       A27R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       A27R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let B27R3 = worksheetFr3.getCell('B27');
       B27R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       B27R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let C27R3 = worksheetFr3.getCell('C27');
       C27R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       C27R3.alignment = { vertical: 'middle', horizontal: 'center' }

       let D27R3 = worksheetFr3.getCell('D27');
       D27R3.font = {
         name: 'Cambria',
         size: 9,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       D27R3.alignment = { vertical: 'middle', horizontal: 'center' }

       
       let E27R3 = worksheetFr3.getCell('E27');
       E27R3.font = {
         name: 'Cambria',
         size: 9,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       E27R3.alignment = { vertical: 'middle', horizontal: 'center' }



      alp3.forEach((item) => {
        worksheetFr3.getCell(`${item}27`).border = {
          bottom: {style:'hair', color: {argb:'000000'}},
          left: item != "F" ? {style:'hair', color: {argb:'000000'}} : {style:'hair', color: {argb:'FFFFFF'}},
          right: {style:'hair', color: {argb:'000000'}},
          top: {style:'hair', color: {argb:'000000'}}
        }
      })

      worksheetFr3.getCell(`E7`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      worksheetFr3.getCell(`E27`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }

      worksheetFr3.getCell(`B1`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        //left: {style:'hair', color: {argb:'000000'}},
        //right: {style:'hair', color: {argb:'000000'}},
        //top: {style:'hair', color: {argb:'000000'}}
      }

      worksheetFr3.getCell(`B2`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        //left: {style:'hair', color: {argb:'000000'}},
        //right: {style:'hair', color: {argb:'000000'}},
        //top: {style:'hair', color: {argb:'000000'}}
      }
      worksheetFr3.getCell(`B3`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        //left: {style:'hair', color: {argb:'000000'}},
        //right: {style:'hair', color: {argb:'000000'}},
        //top: {style:'hair', color: {argb:'000000'}}
      }

      worksheetFr3.getCell(`D3`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        //left: {style:'hair', color: {argb:'000000'}},
        //right: {style:'hair', color: {argb:'000000'}},
        //top: {style:'hair', color: {argb:'000000'}}
      }

      worksheetFr3.getCell(`F3`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        //left: {style:'hair', color: {argb:'000000'}},
        //right: {style:'hair', color: {argb:'000000'}},
        //top: {style:'hair', color: {argb:'000000'}}
      }

      worksheetFr3.getCell(`E2`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        //left: {style:'hair', color: {argb:'000000'}},
        //right: {style:'hair', color: {argb:'000000'}},
        //top: {style:'hair', color: {argb:'000000'}}
      }
      worksheetFr3.getCell(`F2`).border = {
        bottom: {style:'hair', color: {argb:'000000'}},
        //left: {style:'hair', color: {argb:'000000'}},
        //right: {style:'hair', color: {argb:'000000'}},
        //top: {style:'hair', color: {argb:'000000'}}
      }


      for(let i = 28; i <= 45; i++){
       worksheetFr3.getCell(`B${i}`).border =  {
          bottom: i == 45 ? {style:'hair', color: {argb:'000000'}} :{style:'hair', color: {argb:'FFFFFF'}}  ,
          left: {style:'hair', color: {argb:'000000'}},
          right: {style:'hair', color: {argb:'000000'}},
          //top: {style:'hair', color: {argb:'000000'}}
        }
        
      }

      for(let i = 28; i <= 45; i++){
        worksheetFr3.getCell(`C${i}`).border =  {
          bottom: i == 45 ? {style:'hair', color: {argb:'000000'}} :{style:'hair', color: {argb:'FFFFFF'}}  ,
           left: {style:'hair', color: {argb:'000000'}},
           //right: {style:'hair', color: {argb:'000000'}},
           //top: {style:'hair', color: {argb:'000000'}}
         }
         
       }


       for(let i = 28; i <= 45; i++){
        worksheetFr3.getCell(`D${i}`).border =  {
          bottom: i == 45 ? {style:'hair', color: {argb:'000000'}} :{style:'hair', color: {argb:'FFFFFF'}}  ,
           //left: {style:'hair', color: {argb:'000000'}},
           right: {style:'hair', color: {argb:'000000'}},
           //top: {style:'hair', color: {argb:'000000'}}
         }
         
       }

       for(let i = 28; i <= 45; i++){
        worksheetFr3.getCell(`F${i}`).border =  {
           bottom: i == 45 ? {style:'hair', color: {argb:'000000'}} :{style:'hair', color: {argb:'FFFFFF'}}  ,
           //left: {style:'hair', color: {argb:'000000'}},
           right: {style:'hair', color: {argb:'000000'}},
           //top: {style:'hair', color: {argb:'000000'}}
         }
         
       }

       worksheetFr3.getCell('A45').border =  {
        bottom:  {style:'hair', color: {argb:'000000'}},
         left: {style:'hair', color: {argb:'000000'}},
         right: {style:'hair', color: {argb:'000000'}},
         //top: {style:'hair', color: {argb:'000000'}}
       }

       worksheetFr3.getCell('E45').border =  {
        bottom:  {style:'hair', color: {argb:'000000'}},
         left: {style:'hair', color: {argb:'000000'}},
         //right: {style:'hair', color: {argb:'000000'}},
         //top: {style:'hair', color: {argb:'000000'}}
       }


       let A23R3 = worksheetFr3.getCell('A23');
       A23R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         //bold: true,
         color: { argb: '000000' }
       }
       A23R3.alignment = { vertical: 'middle', horizontal: 'left' }

       let A25R3 = worksheetFr3.getCell('A25');
       A25R3.font = {
         name: 'Cambria',
         size: 10,
         //underline: 'single',
         bold: true,
         color: { argb: '000000' }
       }
       A25R3.alignment = { vertical: 'middle', horizontal: 'center' }


     
      /// fin Dirigeants R3






       /// Tableau des Notes R4

       const worksheetR4 = workbook.addWorksheet('Tableau des Notes R4',{views: [{showGridLines: false}]});


  
         worksheetR4.columns = [
           {width: 15},
           {width: 70},
           {width: 10},
           {width: 10}];

         const R4Data = [
             {
              "length": 15
            },
     {
"A1":"NOTES",
"B1":"INTITULES",
"C1":"A",
"D1":"N/A",
"A2":"NOTE 1",
"B2":"DETTES GARANTIES PAR DES SURETES REELLES",	
"A3":"NOTE 2",	
"B3":"INFORMATIONS OBLIGATOIRES",	
"A4":"NOTE 3A",
"B4":"IMMOBILISATION BRUTE",	
"A5":"NOTE 3B",
"B5":"BIENS PRIS EN LOCATION ACQUISITION",		
"A6":"NOTE 3C",
"B6":"IMMOBILISATIONS : AMORTISSEMENTS",		
"A7":"NOTE 3D",	
"B7":"IMMOBILISATIONS : PLUS-VALUES ET MOINS VALUES DE CESSION",		
"A8":"NOTE 3E",
"B8":"INFORMATIONS SUR LES REEVALUATIONS EFFECTUEES PAR L'ENTITE",	
"A9":"NOTE 3F",
"B9":"TABLEAUD'ETALEMENT DES CHARGES IMMOBILISEES",	
"A10":"NOTE 4",
"B10":"IMMOBILISATIONS FINANCIERES",	
"A11":"NOTE 5",
"B11":"ACTIF CIRCULANT HAO",	
"A12":"NOTE 6",
"B12":"STOCKS ET ENCOURS",	
"A13":"NOTE 7",
"B13":"CLIENTS PRODUITS A RECEVOIR",
"A14":"NOTE 8",
"B14":"AUTRES CREANCES",	
"A15":"NOTE 9",
"B15":"TITRES DE PLACEMENT",	
"A16":"NOTE 10",
"B16":"VALEURS A ENCAISSER",	
"A17":"NOTE 11",
"B17":"DISPONIBILITES",		
"A18":"NOTE 12",
"B18":"ECARTS DE CONVERSION",		
"A19":"NOTE 13",
"B19":"CAPITAL : VALEUR NOMINALE DES ACTIONS OU PARTS",	
"A20":"NOTE 14",
"B20":"PRIMES ET RESERVES",		
"A21":"NOTE 15A",	
"B21":"SUBVENTIONS ET PROVISIONS REGLEMENTEES",		
"A22":"NOTE 15B",	
"B22":"AUTRES FONDS PROPRES",		
"A23":"NOTE 16A",
"B23":"DETTES FINANCIERES ET RESSOURCES ASSIMILEES",		
"A24":"NOTE 16B",
"B24":"ENGAGEMENTS DE RETRAITE ET AVANTAGES  ASSIMILES (METHODE ACTURIELLE)",		
"A25":"NOTE 16B bis",
"B25":"ENGAGEMENTS DE RETRAITE ET AVANTAGES  ASSIMILES (METHODE ACTURIELLE)",	
"A26":"NOTE 16C",	
"B26":"ACTIFS ET PASSIFS EVENTUELS",		
"A27":"NOTE 17",
"B27":"FOURNISSEURS D'EXPLOITATION",
"A28":"NOTE 18",
"B28":"DETTES FISCALES ET SOCIALES",		
"A29":"NOTE 19",
"B29":"AUTRES DETTES ET PROVISIONS POUR RISQUES A COURT TERME",		
"A30":"NOTE 20",
"B30":"BANQUES, CREDIT D'ESCOMPTE ET TRESORERIE",	
"A31":"NOTE 21",
"B31":"CHIFFRE D'AFFAIRES ET AUTRES PRODUITS",	
"A32":"NOTE 22",	
"B32":"ACHATS",	
"A33":"NOTE 23",	
"B33":"TRANSPORTS",		
"A34":"NOTE 24",
"B34":"SERVICES EXTERIEURS",		
"A35":"NOTE 25",
"B35":"IMPOTS ET TAXES",	
"A36":"NOTE 26",	
"B36":"AUTRES CHARGES",	
"A37":"NOTE 27A",
"B37":"CHARGES DE PERSONNEL",	
"A38":"NOTE 27B",
"B38":"EFFECTIFS, MASSE SALARIALE ET PERSONNEL EXTERIEUR",		
"A39":"NOTE 28",	
"B39":"PROVISIONS ET DEPRECIATIONS INSCRITES AU BILAN",		
"A40":"NOTE 29",
"B40":"CHARGES ET REVENUS FINANCIERS",	
"A41":"NOTE 30",
"B41":"AUTRES CHARGES ET PRODUITS HAO",	
"A42":"NOTE 31",	
"B42":"REPARTITION DU RESULTAT ET AUTRES ELEMENTS CARACTERISTIQUES DES CINQ DERNIERES ANNEES",	
"A43":"NOTE 32",	
"B43":"PRODUCTION DE L'EXERCICE",	
"A44":"NOTE 33",
"B44":"ACHATS DESTINES A LA PRODUCTION",	
"A45":"NOTE 34",
"B45":"FICHE DE SYNTHESE DES PRINCIPAUX INDICATEURS FINANCIERS",	
"A46":"NOTE 35",
"B46":"LISTE DES INFORMATIONS SOCIALES, ENVIRONNEMENTALES ET SOCIALES A FOURNIR",		
"A47":"NOTE 36",
"B47":"TABLES DES CODES",
            }
         ];

         const R4alph =  [
          'A',
          'B',
          'C',
          'D',
          ];
  


       for(let i = 1 ; i <= 47; i++){

        R4alph.forEach((a) => {
          worksheetR4.getCell(`${a}${i}`).border =  {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            top: {style:'hair', color: {argb:'000000'}}
          }

          worksheetR4.getRow(i).height = 25;

        })}



         for(let i = 0; i <= Object.keys(R4Data).length; i++ ){

          for(var v in R4Data[1]){
      
          
               console.log(`mon ${v}`)
                let cell = worksheetR4.getCell(v);
               cell.value = R4Data[1][v];
                cell.font = {
                  name: 'Cambria',
                  size: 9,
                  //underline: 'single',
                 // bold: true,
                  color: { argb: '000000' }
                }
          
                cell.alignment = { vertical: 'middle', horizontal: 'justify'}
      
              
              
        
          }
      
        }



        


         


       /// Tableau des Notes R4


        
       /// BILAN PAYSAGE

       const worksheetBP = workbook.addWorksheet('BILAN PAYSAGE',{views: [{showGridLines: true}]});


       worksheetBP.mergeCells('A1','A2');
       worksheetBP.mergeCells('I1','I2');
       worksheetBP.mergeCells('D1','F1');


       


       worksheetBP.getRow(1).height = 30;
       worksheetBP.getRow(2).height = 20;


     let  BPalph: string[] = [
        'A',
        'B',
        'C',
        'D',
        'E',
        'F',
        'G',
        'H',
        'I',
        'J',
        'K',
        'L'
        ];


        for(let i = 1 ; i < 32; i++){

        BPalph.forEach((a) => {

          worksheetBP.getCell(`${a}${i}`).border =  {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            top: {style:'hair', color: {argb:'000000'}}
          }

        })
          
        if(i >= 3){
         worksheetBP.getRow(i).height = 28;
        }


        }




        

      worksheetBP.columns = 
       [{width: 5},
        {width: 50},
        {width: 5},
        {width: 15},
        {width: 15},
        {width: 15},
        {width: 15},
        {width: 5},
        {width: 40},
        {width: 5},
        {width: 15},
        {width: 15}];



     const BPData = [
       {
          "lignes": 31
       },
       {
        "A1":"REF",	
        "B1":"ACTIF",
        "C1":"Note",		
        "E1":"Exercice au 31/12/2019",
        "G1":"Exercice au 31/12/2018",
        "I1":"REF	PASSIF",
        "J1":"Note",	
        "K1":"Exercice au 31/12/2019",
        "L1":"Exercice au 31/12/2018",
        "D2":"Brut",	
        "E2":"Amorts et Dépréc.",
        "F2":"Net",	
        "G2":"Net",				
        "K2":"Net",	
        "L2":"Net",
        "A3":"AD",	
        "B3":"IMMOBILISATIONS INCORPORELLES",	
        "C3":"3",					
        "H3":"CA",	
        "I3":"CAPITAL",
        "J3":"13",	 
        "K3":"56 000 000",   	
        "A4":"AE",	
        "B4":"Frais de développement et de prospection",
        "H4":"CB",
        "I4":"Apporteurs capital non appelé",
        "J4":"13",		
        "A5":"AF",
        "B5":"Brevets, licences, logiciels et droits similaires",
        "H5":"CD",
        "I5":"Primes liées au capital social",
        "J5":"14",
        "A6":"AG",
        "B6":"Fonds commercial et droit au bail",
        "H6":"CE",
        "I6":"Ecarts de réévaluation",
        "J6":"3e",
        "A7":"AH",	
        "B7":"Autres immobilisations incorporelles",
        "H7":"CF",
        "I7":"Réserves indisponibles",
        "J7":" 14",
        "A8":"AI",	
        "B8":"IMMOBILISATIONS CORPORELLES",
        "C8":"3",				
        "H8":"CG",	
        "I8":"Réserves libres",
        "J8":"14",	
        "A9":"AJ",	
        "B9":"Terrains (1) (1) dont placement en net   ………../………… ",
        "H9":"CH",
        "I9":"Report à nouveau",
        "J9":"14",
        "A10":"AK",	
        "B10":"Bâtiments (1) dont placement en net   ………../………… ",
        "H10":"CJ",	
        "I10":"Résutat net de l'exercice (bénéfice + ou perte -)",
        "A11":"AL",
        "B11":"Aménagements, agencements et installations",
        "H11":"CL",
        "I11":"Subventions d'investissement",
        "J11":"15",
        "A12":"AM",
        "B12":"Matériel, mobilier et actifs biologiques",
        "H12":"CM",	
        "I12":"Provisions réglementées et fonds assimilés",
        "J12":"15",	
        "A13":"AN",
        "B13":"Matériel de transport",
        "H13":"CP",
        "I13":"TOTAL CAPITAUX PROPRES ET RESSOURCES ASSIMILEES",
        "A14":"AP",
        "B14":"Avances & acomptes versés sur immobilisations",
        "C14":"3",					
        "H14":"DA",
        "I14":"Emprunts et dettes financières diverses",
        "J14":"16",
        "A15":"AQ",
        "B15":"IMMOBILISATIONS FINANCIERES",
        "C15":"4",					
        "H15":"DB",
        "I15":"Dettes de location acquisition",
        "J15":"16",
        "A16":"AR",
        "B16":"Titres de participation",
        "H16":"DC",
        "I16":"Provisions pour risques et charges",
        "J16":"16",
        "A17":"AS",
        "B17":"Autres immobilisations financières",
        "H17":"DD",
        "I17":"TOTAL DETTES FINANCIERES ET RESSOURCES ASSIMILEES",
        "A18":"AZ",
        "B18":"TOTAL ACTIF IMMOBILISE",
        "H18":"DF",	
        "I18":"TOTAL RESSOURCES STABLES",
        "A19":"BA",
        "B19":"ACTIF CIRCULANT H.A.O.",
        "C19":"5",				
        "H19":"DH",	
        "I19":"Dettes circulantes HAO",	
        "J19":"5",	
        "A20":"BB",	
        "B20":"STOCKS ET ENCOURS",
        "C20":"6",
        "H20":"DI",	
        "I20":"Clients, avances reçues",
        "J20":"7",	
        "A21":"BG",	
        "B21":"CREANCES ET EMPLOIS ASSIMILES",
        "H21":"DJ",	
        "I21":"Fournisseurs d'exploitation",
        "J21":"17",
        "A22":"BH",
        "B22":"Fournisseurs, avances versées",
        "C22":"17",
        "H22":" DK",
        "I22":"Dettes fiscales et sociales",
        "J22":"18",
        "A23":"BI",
        "B23":"Clients",
        "C23":"7",
        "H23":"DM",
        "I23":"Autres dettes",
        "J23":"19",
        "A24":"BJ",
        "B24":"Autres créances",
        "C24":"8",
        "H24":"DN",
        "I24":"Provisions pour risques à court terme",
        "J24":"19",
        "A25":"BK",
        "B25":"TOTAL ACTIF CIRCULANT",
        "H25":"DP",	
        "I25":"TOTAL PASSIF CIRCULANT",
        "A26":"BQ",	
        "B26":"Titres de placement",
        "C26":"9",
        "A27":"BR",
        "B27":"Valeurs à encaisser",
        "C27":"10",
        "H27":"DQ",
        "I27":"Banques, crédit d'escompte",
        "J27":"20",
        "A28":"BS",	
        "B28":"Banques, chèques postaux, caisse et assimilés",
        "C28":"11",
        "H28":"DR",
        "I28":"Banques, établissements financiers et crédits de trésorerie",
        "J28":"20",
        "A29":"BT",
        "B29":"TOTAL TRESORERIE - ACTIF",
        "H29":"DT",	
        "I29":"TOTAL TRESORERIE - PASSIF",
        "A30":"BU",	
        "B30":"Ecarts de conversion - Actif",	
        "C30":"12",
        "H30":"DV",	
        "I30":"Ecarts de conversion - Passif",	
        "J30":"12",	
        "A31":"BZ",
        "B31":"TOTAL GENERAL",						
        "H31":"DZ",	
        "I31":"TOTAL GENERAL"		
     }
     ];

      let BPA1 = worksheetBP.getCell('A1');
      BPA1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPA1.alignment = { vertical: 'middle', horizontal: 'center' }

      let BPB1 = worksheetBP.getCell('B1');
      BPB1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPB1.alignment = { vertical: 'middle', horizontal: 'center' }

      let BPC1 = worksheetBP.getCell('C1');
      BPC1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPC1.alignment = { vertical: 'bottom', horizontal: 'center' }


      let BPE1 = worksheetBP.getCell('E1');
      BPE1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPE1.alignment = { vertical: 'bottom', horizontal: 'center' }

      let BPG1 = worksheetBP.getCell('G1');
      BPG1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPG1.alignment = { vertical: 'bottom', horizontal: 'center' }


      let BPH1 = worksheetBP.getCell('H1');
      BPH1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPH1.alignment = { vertical: 'bottom', horizontal: 'center' }

      let BPI1 = worksheetBP.getCell('I1');
      BPI1.value = "PASSIF";
      BPI1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPI1.alignment = { vertical: 'middle', horizontal: 'right' }

      let BPL1 = worksheetBP.getCell('L1');
      BPL1.value = "Exercice au 31/12/2018";
      BPL1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPL1.alignment = { vertical: 'middle', horizontal: 'justify' }


      let BPJ1 = worksheetBP.getCell('J1');
      BPJ1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPJ1.alignment = { vertical: 'middle', horizontal: 'justify' }

      let BPK1 = worksheetBP.getCell('K1');
      BPK1.font = {
        name: 'Cambria',
        size: 9,
        //underline: 'single',
       // bold: true,
        color: { argb: '000000' }
      }

      BPK1.alignment = { vertical: 'middle', horizontal: 'justify'}


      for(let i = 0; i <= Object.keys(BPData).length; i++ ){

        for(var v in BPData[1]){
    
        
             console.log(`mon ${v}`)
              let cell = worksheetBP.getCell(v);
             cell.value = BPData[1][v];
              cell.font = {
                name: 'Cambria',
                size: 9,
                //underline: 'single',
                bold: v == "A3" || 
                      v == "B3" || 
                      v == "C3" ||
                      v == "A8" || 
                      v == "B8" || 
                      v == "C8" ||
                      v == "A15" || 
                      v == "B15" || 
                      v == "C15" ||
                      v == "A18" || 
                      v == "B18" || 
                      v == "C18" ||
                      v == "A25" || 
                      v == "B25" || 
                      v == "C25" ||
                      v == "A29" || 
                      v == "B29" || 
                      v == "C29" ||
                      v == "A31" || 
                      v == "B31" || 
                      v == "C31" ||
                      v == "A3" || 
                      v == "B3" ||
                      v == "I4"  || 
                      v == "J4"  ||
                      v == "H15" || 
                      v == "I15" ||
                      v == "J15" ||
                      v == "H13" || 
                      v == "I13" ||
                      v == "I4"  ||
                      v == "I18"  ||
                      v == "H18"  ||
                      v == "I24"  ||
                      v == "H24"  ||
                      v == "I25"  ||
                      v == "H25"  ||
                      v == "I17"  ||
                      v == "H17"  ||
                      v == "I29"  ||
                      v == "H29"  ||
                      v == "I31"  ||
                      v == "H31"  ||
                      v == "C3"? true : false,
                color: 
                v == "H4"  || 
                v == "I4"  || 
                v == "J4"  ||
                v == "H15" || 
                v == "I15" ||
                v == "J15" ||
                v == "H24" || 
                v == "I24" ||
                v == "J24" ? { argb: 'FF0000' } : { argb: '000000' }
              }
        
              cell.alignment = { vertical: 'middle', horizontal: 'justify'}
    
            
            
      
        }
    
      }

  /// COMPTE DE RESULTAT

  const worksheetCR = workbook.addWorksheet('COMPTE DE RESULTAT',{views: [{showGridLines: true}]});


  worksheetCR.columns = 
  [{width: 5},
   {width: 50},
   {width: 5},
   {width: 10},
   {width: 15},
   {width: 15},
   ];

   const CRData = [
     {
       "lignes": 43

     },
     {
       "A1": "REF",
       "B1": "LIBELLE",
       "C1": "",
       "D1": "NOTE",
       "E1": "31/12/2019",
       "F1": "31/12/2018",
       "A2": "TA",
       "B2": "Ventes de marchandises A",
       "C2": "+",
       "D2": "21",
       "A3": "RA",
       "B3": "Achats de marchandises",
       "C3": "-",
       "D3": "22",
       "A4": "RB",
       "B4": "Variation de stocks",
       "C4": "-/+",
       "D4": "6",
       "A5": "XA",
       "B5": "MARGE BRUTE SUR MARCHANDISES (somme TA à RB)",
       "A6": "TB",
       "B6": "Ventes de produits fabriqués B",
       "C6": "+",
       "D6": "21",
       "A7": "TC",
       "B7": "Travaux, services vendus         C",
       "C7": "+",
       "D7": "21",
       "A8": "TD",
       "B8": "Produits accessoires       D",
       "C8": "+",
       "D8": "21",
       "A9": "XB",
       "B9": "CHIFFRE D'AFFAIRES (A + B + C + D)",
       "D9": "75000000",
       "A10": "TE",
       "B10": "Production stockée (ou destockage)",
       "C10": "-/+",
       "D10": "6",
       "A11": "TF",
       "B11": "Production immobilisée",
       "D11": "21",
       "A12": "TG",
       "B12": "Subventions d'exploitation",
       "D12": "21",
       "A13": "TH",
       "B13": "Autres produits",
       "C13": "+",
       "D13": "21",
       "A14": "TI",
       "B14": "Transferts de charges",
       "C14": "+",
       "D14": "12",
       "A15": "RC",
       "B15": "Achats de matières premières et fournitures liées",
       "C15": "-",
       "D15": "12",
       "A16": "RD",
       "B16": "Variation de stocks de stocks de matières premières et fournitures liées",
       "C16": "-/+",
       "D16": "6",
       "A17": "RE",
       "B17": "Autres achats",
       "C17": "-",
       "D17": "22",
       "A18": "RF",
       "B18": "Variation de stocks d'autres approvisionnements",
       "C18": "-/+",
       "D18": "6",
       "A19": "RG",
       "B19": "Transports",
       "C19": "-",
       "D19": "23",
       "A20": "RH",
       "B20": "Services extérieurs",
       "C20": "-",
       "D20": "24",
       "A21": "RI",
       "B21": "Impôts et taxes",
       "C21": "-",
       "D21": "25",
       "A22": "RJ",
       "B22": "Autres charges",
       "C22": "-",
       "D22": "26",
       "A23": "XC",
       "B23": "VALEUR AJOUTEE (XB + RA + RB) + (somme TE à RJ)",
       "A24": "RK",
       "B24": "Charges de personnel",
       "C24": "-",
       "D24": "27",
       "A25": "XD",
       "B25": "EXCEDENT BRUT D'EXPLOITATION (XC + RK)",
       "A26": "TJ",
       "B26": "Reprises d'amortissements",
       "C26": "+",
       "D26": "28",
       "A27": "RL",
       "B27": "Dotations aux amortissements, aux provisions et dépréciations",
       "C27": "-",
       "D27": "3C & 28",
       "A28": "XE",
       "B28": "RESULTAT D'EXPLOITATION (XD + TJ + RL)",
       "A29": "TK",
       "B29": "Revenus financiers et assimilés",
       "C29": "+",
       "D29": "29",
       "A30": "TL",
       "B30": "Reprises de provisions et dépréciations financières",
       "C30": "+",
       "D30": "28",
       "A31": "TM",
       "B31": "Tranfert de charges financières",
       "C31": "+",
       "D31": "12",
       "A32": "RM",
       "B32": "Frais financiers et charges assimilées",
       "C32": "+",
       "D32": "12",
       "A33": "RN",
       "B33": "Dotations aux provisions et aux dépréciations financières",
       "C33": "-",
       "D33": "3C & 28",
       "A34": "XF",
       "B34": "RESULTAT FINANCIER (somme TK à RN)",
       "A35": "XG",
       "B35": "RESULTAT DES ACTIVITES ORDINAIRES (XE + XF)",
       "A36": "TN",
       "B36": "Produits des cessions d'immobilisations",
       "C36": "+",
       "D36": "3D",
       "A37": "TO",
       "B37": "Autres produits H.A.O.",
       "C37": "+",
       "D37": "30",
       "A38": "RO",
       "B38": "Valeurs comptables des cessions d'immobilisations",
       "C38": "-",
       "D38": "3D",
       "A39": "RP",
       "B39": "Charges H.A.O.",
       "C39": "-",
       "D39": "30",
       "A40": "XH",
       "B40": "RESULTAT HORS ACTIVITES ORDINAIRES (somme TN à RP)",
       "A41": "RQ",
       "B41": "Participations des travailleurs",
       "C41": "-",
       "D41": "30",
       "A42": "RS",
       "B42": "Impôts sur le résultat",
       "C42": "-",
       "A43": "XI",
       "B43": "RESULTAT NET (XG + XH + RQ + RS)",
       "E43": "2500000000"

     }
   ];


      

     const CRalph =  [
        'A',
        'B',
        'C',
        'D',
        'E',
        'F',
        ];

       

       for(let i = 1 ; i <= CRData[0]['lignes']; i++){

        CRalph.forEach((a) => {
          worksheetCR.getCell(`${a}${i}`).border =  {
            bottom: {style:'hair', color: {argb:'000000'}},
            left: {style:'hair', color: {argb:'000000'}},
            right: {style:'hair', color: {argb:'000000'}},
            top: {style:'hair', color: {argb:'000000'}}
          }

        })}


        for(let i = 0; i <= Object.keys(CRData).length; i++ ){

          for(var v in CRData[1]){
      
          
               console.log(`mon ${v}`)
                let cell = worksheetCR.getCell(v);
               cell.value = CRData[1][v];
                cell.font = {
                  name: 'Cambria',
                  size: 9,
                  //underline: 'single',
                 // bold: true,
                  color: { argb: '000000' }
                }
          
                cell.alignment = { vertical: 'middle', horizontal: 'justify'}
      
              
              
        
          }
      
        }



  /// FLUX DE TRESORERIE

  const worksheetFR = workbook.addWorksheet('FLUX DE TRESORERIE',{views: [{showGridLines: true}]});


  worksheetFR.columns = 
  [{width: 5},
   {width: 70},
   {width: 5},
   {width: 13},
   {width: 20},
   {width: 20},
   ];

   const FRData = [
     {
       "A1": "FL",
       "B1": "+ Subventions d'investissements reçues",
       "B2": "Trésorerie nette au 1er janvier (Trésorerie actif N-1 - Trésorerie passif N-1)",
       "A2": "ZA",
       "C2": "2",
       "B3": "Flux  de trésorerie provenant des activités opérationnelles",
       "A4": "FA",
       "B4":"Capacité d'Autofinancement Globale (CAFG)",
       "A5": "FA",
       "B5": "- Variation de l'actif circulant HAO (1)",
       "A6": "FC",
       "B6": "- Variation des stocks",
       "A7": "FD",
       "B7": "- Variation des créances",
       "A8": "FE",
       "B8": "+ Variation du passif circulant (1)",
       "B9": "Variation du BF lié aux activités opérationnelles (FB + FC + FD + FE) : ",
       "A10": "ZB",
       "B10": "Flux de trésorerie provenant des activités opérationnelles (somme FA à FE)",
       "C10": "B",
       "B11": "Flux  de trésorerie provenant des activités d'investissement",
       "A12": "FF",
       "B12": "- Décaissements liés aux acquisitions d'immobilisations incorporelles",
       "A13": "FG",
       "B13": "- Décaissements liés aux acquisitions d'immobilisations corporelles",
       "A14": "FH",
       "B14": "- Décaissements liés aux acquisitions d'immobilisations financières",
       "A15": "FI",
       "B15": "+ Encaissements liés aux cessions d'immobilisations incorporelles et corporelles",
       "A16": "FJ",
       "B16": "+ Encaissements liés aux cessions d'immobilisations financières",
       "A17": "ZC",
       "B17": "Flux  de trésorerie provenant des activités d'investissement (somme FF à FJ)",
       "C17": "C",
       "A18": "Flux  de trésorerie provenant du financement par les capitaux propres",
       "A19": "FK",
       "B19": "+ Augmentations de capital par apports nouveaux",
       "A20": "FL",
       "B20": "+ Subventions d'investissements reçues",
       "A21": "FM",
       "B21": "- Prélèvements sur le capital",
       "A22": "FN",
       "B22": "- Dividendes versés",
       "D22": "D",
       "A23": "ZD",
       "B23": "Flux  de trésorerie provenant du financement par les capitaux propres (somme FK à FN)",
       "B24": "Flux  de trésorerie provenant du financement par les capitaux étrangers",
       "A25": "FO",
       "B25": "+ Emprunts",
       "A26": "FP",
       "B26": "+ Autres dettes financières",
       "A27":"FQ",
       "B27": "- Remboursements des emprunts et autres dettes financières",
       "A28": "ZE",
       "B28": "Flux  de trésorerie provenant du financement par les capitaux étrangers (somme FO à FQ)",
       "C28": "E",
       "A29": "ZF",
       "B29": "Flux  de trésorerie provenant des capitaux étrangers (D + E)",
       "C29": "F",
       "A30": "ZG",
       "B30": "VARIATION DE LA TRESORERIE NETTE DE LA PERIODE (B + C + F)",
       "C30": "G",
       "A31": "XI",
       "B31": "Trésorerie nette au 31 décembre (G + A) Contrôle : Trésorerie actif N - Trésorerie passif N",
       "C31": "H",
       "B32": "Contrôle"








     }
   ];


      

     const FRalph =  [
        'A',
        'B',
        'C',
        'D',
        'E',
        'F',
        ];



       for(let i = 1 ; i <= 32; i++){
          FRalph.forEach((a) => {
            worksheetFR .getCell(`${a}${i}`).border =  {
              bottom: {style:'hair', color: {argb:'000000'}},
              left: {style:'hair', color: {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              top: {style:'hair', color: {argb:'000000'}}
            }
  
          })
           }
        


        for(let i = 0; i <= CRData.length; i++ ){

          FRData.forEach((v,k) =>{
      
          
              for(var d in v)  {
               console.log(d)
                let cell = worksheetFR.getCell(d);
                cell.value = v[d];
                cell.font = {
                  name: 'Cambria',
                  size: 9,
                  //underline: 'single',
                 // bold: true,
                  color: { argb: '000000' }
                }
          
                cell.alignment = { vertical: 'middle', horizontal: 'justify'}
      
              }
              
        
          })
      
        }



    /// Note 1

  const worksheetN1 = workbook.addWorksheet('Note 1',{views: [{showGridLines: true}]});


  worksheetN1.mergeCells('A1','F2');


  worksheetN1.columns = 
  [{width: 50},
   {width: 8},
   {width: 20},
   {width: 20},
   {width: 20},
   {width: 20},
   ];

   const N1alph =  [
    'A',
    'B',
    'C',
    'D',
    'E',
    'F',
    ];

   const N1Data = [
     {
       "A1": "NOTE 1 : DETTES GARANTIES PAR DES SURETES REELLES",
       "A3": "Désignation entité :",
       "D3": "Excercice clos le :",
       "E3": "31/12/2019",
       "A5": "Numéro d'identification :",
       "B5": "0049",
       "D4": "Durée en mois :",
       "E4": "12",
       "A7": "LIBELLES",
       "B7": "Note",
       "C7": "Montant brut",
       "D8": "Hypothèques",
       "D7": "SURETES REELLES",
       "F8": "Gages",
       "E8": "Nantissements",
       "F9": "Autres",
       "A10": "Dettes financières et ressources assimilées :",
       "A11": "Emprunts obligataires convertibles",
       "A12": "Autres emprunts obligataires",
       "A13": "Emprunts et dette des établissements de crédit",
       "A14": "Autres dettes financières",
       "A15": "SOUS TOTAL (1)",
       "C15":"0",
       "D15": "0",
       "E15": "0",
       "F15": "0",
       "A17": "Dettes de location-acquisition :",
       "A18": "Dettes de crédit-bail immobilier",
       "A19": "Dettes de crédit-bail mobilier",
       "A20": "Dettes sur contrats de location-vente",
       "A21": "Dettes sur contrats de location-acquisition",
       "A22": "SOUS TOTAL (1)",
       "C22":"0",
       "D22": "0",
       "E22": "0",
       "F22": "0",
       "A24": "Dettes du passif circulant :",
       "A25": "Fournisseurs",
       "A26": "Clients",
       "A27": "Personnel",
       "A28": "Sécurité sociale et organismes sociaux",
       "A29": "Etat",
       "A30": "Organismes internationaux",
       "A31": "Associéss et groupe",
       "A32": "Crédits divers",
       "A33": "SOUS TOTAL (1)",
       "C33":"0",
       "D33": "0",
       "E33": "0",
       "F33": "0",
       "A34": "TOTAL (1) + (2) + (3)",
       "C34":"0",
       "D34": "0",
       "E34": "0",
       "F34": "0",
       "A35": "ENGAGEMENTS FINANCIERS",
       "E35": "Engagements donnés",
       "F35": "Engagements reçus",
       "A36": "Engagements consentis à des entités liées",
       "A37":"Primes de remboursement non échus",
       "A38": "Avals, cautions, garanties",
       "A39": "Hypothèques, nantissements, gages, autres",
       "A40": "Effets escomptés non échus",
       "A41": "Créances commerciales et professionnelles cédées",
       "A42": "Abandon de créances conditionnelles",
       "A43": "TOTAL",
       "D43": "0",
       "E43": "0",
       "A44": "Commentaire :",
       "A45":"· Indiquer la raison d’être des suretés"



     }
   ];


      

     



       for(let i = 1 ; i <= 45; i++){
       

          N1alph.forEach((a) => {
            if(`${a}${i}` != "A8"){
              worksheetN1 .getCell(`${a}${i}`).border =  {
                bottom: {style:'hair', color: {argb:'000000'}},
                left: {style:'hair', color: {argb:'000000'}},
                right: {style:'hair', color: {argb:'000000'}},
                top: {style:'hair', color: {argb:'000000'}}
              }
            }
            
  
          })

        
        }


        for(let i = 0; i <= N1Data.length; i++ ){

          N1Data.forEach((v,k) =>{
      
          
              for(var d in v)  {
               console.log(d)
                let cell = worksheetN1.getCell(d);
                cell.value = v[d];
                cell.font = {
                  name: 'Cambria',
                  size: 9,
                  //underline: 'single',
                 bold: d  === "A1"? true : false,
                  color: { argb: '000000' }
                }
          
                cell.alignment = { vertical: 'middle', horizontal: d === "A1"? 'center' : 'justify'}
      
              }
              
        
          })
      
        }





         /// Note 2

  const worksheetN2 = workbook.addWorksheet('Note 2',{views: [{showGridLines: true}]});
  worksheetN2.mergeCells('A1','F2');
  worksheetN2.columns = 
  [{width: 50},
   {width: 8},
   {width: 20},
   {width: 20},
   {width: 20},
   {width: 20},
   ];

   const N2alph =  [
    'A',
    'B',
    'C',
    'D',
    'E',
    'F',
    ];

   const N2Data = [
     {
      "A1":"NOTE 2 : INFORMATIONS OBLIGATOIRES",		
      "A3":"Désignation entité :",		
      "C3":"Excercice clos le :",
      "D3":"31/12/2019",
      "A4":"0",	
      "C4":"Durée en mois :",
      "D4":"12",
      "A5":"Numéro d'identification :",	
      "B5":"004942901",		
      "A7":"A - DECLARATION DE CONFORMITE AU SYSCOHADA",			
      "A8":"Les états financiers sont établis en conformité avec le systéme comptable OHADA et l'acte uniforme relatif au droit comptable et à l'information financiére",		
      "A9":"B - REGLES ET METHODES COMPTABLES",			
      "A10":"les états financiers ont été confectionnés dans le respect des postulats, des conventions et des régles d'évaluation édictés par le SYSCOHADA et l'Acte Uniforme",			
      "A11":"C - DEROGATION AUX POSTULATS ET CONVENTIONS COMPTABLES",			
      "A12":"Respect de tous les postulats et conventions comptables sans aucune dérogation",			
      "A13":"D - INFORMATIONS COMPLEMENTAIRES RELATIVES AU BILAN, AU COMPTE DE RESULTAT ET AU TABLEAU DES FLUX DE TRESORERIE",			
      "A14":"Pas d'informations complémentaires relatives aux autres états financiers"
}
   ];

       for(let i = 1 ; i <= 14; i++){
          N2alph.forEach((a) => {
            if(`${a}${i}` != "A8"){
              worksheetN2 .getCell(`${a}${i}`).border =  {
                bottom: {style:'hair', color: {argb:'000000'}},
                left: {style:'hair', color: {argb:'000000'}},
                right: {style:'hair', color: {argb:'000000'}},
                top: {style:'hair', color: {argb:'000000'}}
              }
            }})}
        for(let i = 0; i <= N2Data.length; i++ ){
          N2Data.forEach((v,k) =>{
              for(var d in v)  {
               console.log(d)
                let cell = worksheetN2.getCell(d);
                cell.value = v[d];
                cell.font = {
                  name: 'Cambria',
                  size: 9,
                  //underline: 'single',
                 bold: d  === "A1"? true : false,
                  color: { argb: '000000' }
                }
                cell.alignment = { vertical: 'middle', horizontal: d === "A1"? 'center' : 'justify'}
              } })
      
        }












/// Note 3

const worksheetN3 = workbook.addWorksheet('Note 3',{views: [{showGridLines: true}]});
worksheetN3.mergeCells('A1','F2');
worksheetN3.columns = 
[{width: 50},
 {width: 30},
 {width: 20},
 {width: 20},
 {width: 20},
 {width: 20},
 {width: 20},

 ];

 const N3alph =  [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H'
  ];



 const N3Data = [
   {
    "A1":"NOTE 3A : IMMOBILISATION BRUTE",									
    "A3":"Désignation entité :",				
    "E3":"Excercice clos le :",	
    "F3":"31/12/2019",		
    "A4":"0",		
    "E4":"Durée en mois :",
    "F4":"12",		
    "A5":"Numéro d'identification :",	
    "B5":"004942901",            
    "A7":"SITUATIONS ET MOUVEMENTS",
    "B7":"Montant brut à l'ouverture de l'exercice",
    "C7":"Acquisitions Apports Créations",
    "D7":"Virements de poste à poste",
    "E7":"Suite à une réévaluation pratiquée au cours de l'exercice",
    "F7":"Cessions Scissions Hors service",	
    "G7":"Virements de poste à poste",	
    "H7":"Montant brut à la clôture de l'exercice",
    "A10":"IMMOBILISATIONS INCORPORELLES",						
    "A11":"Frais de développement et de prospection",							
    "A12":"Brevets, licences, logiciels et droits similaires",						
    "A13":"Fonds commercial et droit au bail",							
    "A14":"Autres immobilisations incorporelles",							
    "A15":"IMMOBILISATIONS CORPORELLES",						
    "A16":"Terrains hors immeuble de placement",							
    "A17":"Terrains immeuble de placement",							
    "A18":"Bâtiments hors immeuble de placement",							
    "A19":"Bâtiments immeuble de placement",							
    "A20":"Aménagements, agencements et installations",							
    "A21":"Matériel, mobilier et actifs biologiques",							
    "A22":"Matériel de transport",							
    "A23":"AVANCES ET ACOMPTES VERSES SUR IMMOBILISATIONS",							
    "A24":"Immobilisations incorporelles",							
    "A25":"Immoblisations corporelles",							
    "A26":"IMMOBILISATIONS FINANCIERES",							
    "A27":"Titres de placement",							
    "A28":"Autres immobilisations financières",							
    "A29":"TOTAL GENERAL",					
    "A30":"Commentaire :",							
    "A31":"· Toute variation significative doit être commentée",						
    "A32":"· Détailler les éléments constitutifs du fonds commercial et indiquer la date d'acquisition",
    "A33":"· Pour l'immobilisation incorporelle relative à la concession faire un descriptif de l'accord",							
    "A34":"· Indiquer :",							
    "A35":"- la nature de la créance ;",						
    "A36":"- la durée de la concession ;",						
    "A37":"- l'échéance ;",							
    "A38":"· Indiquer les créances du groupe avec nature et date d'échéance",
    "A39":"· Pour les banques, DAT indiquer le nom de labanque, le montant et la date d'échéance"	
}
 ];

     for(let i = 1 ; i <= 39; i++){
        N3alph.forEach((a) => {
          if(`${a}${i}` != "A8"){
            worksheetN3.getCell(`${a}${i}`).border =  {
              bottom: {style:'hair', color: {argb:'000000'}},
              left: {style:'hair', color: {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              top: {style:'hair', color: {argb:'000000'}}
            }
          }})}
      for(let i = 0; i <= N3Data.length; i++ ){
        N3Data.forEach((v,k) =>{
            for(var d in v)  {
             console.log(d)
              let cell = worksheetN3.getCell(d);
              cell.value = v[d];
              cell.font = {
                name: 'Cambria',
                size: 9,
                //underline: 'single',
               bold: d  === "A1"? true : false,
                color: { argb: '000000' }
              }
              cell.alignment = { vertical: 'middle', horizontal: d === "A1"? 'center' : 'justify'}
            } })
    
      }




      ///Note 3B  

const worksheetNB3 = workbook.addWorksheet('Note 3B',{views: [{showGridLines: true}]});
worksheetNB3 .mergeCells('A1','F2');
worksheetNB3 .columns = 
[{width: 40},
 {width: 15},
 {width: 15},
 {width: 15},
 {width: 15},
 {width: 15},
 {width: 15},
 {width: 15}];

 const NB3alph =  [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I'
  ];

 const NB3Data = [
   {
    "A1":"NOTE 3B : BIENS PRIS EN LOCATION ACQUISITION",								
    "A3":"Désignation entité :",					
    "G3":"Excercice clos le :	31/12/2019",
    "A4":"0",						
    "G4":"Durée en mois :",	
    "H4":"12",	
    "A5":"Numéro d'identification :",
    "B5":"004942901",						
    "A7":"SITUATIONS ET MOUVEMENTS",
    "B7":"NATURE DU CONTRAT      (I; M; A)",
    "C7":"A",
    "D7":"AUGMENTATIONS  B",		
    "G7":"DIMINUTIONS  C",		
    "I7":"D = A +B - C",
    "B9":"Montant brut à l'ouverture de l'exercice",	
    "C8":"Acquisitions Apports Créations",	
    "D8":"Virements de poste à poste",	
    "E8":"Suite à une réévaluation pratiquée au cours de l'exercice",
    "F8":"Cessions Scissions Hors service",	
    "G8":"Virements de poste à poste",	
    "H8":"Montant brut à l'ouverture de l'exercice",
    "I8":"[1]",							
    "A9":"Brevets, licences, logiciels et droits similaires",							
    "A10":"Fonds commercial et droit au bail",							
    "A11":"Autres immobilisations incorporelles",							
    "A12":"SOUS TOTAL : IMMOBILISATIONS INCORPORELLES",								
    "A13":"Terrains",							
    "A14":"Bâtiments",							
    "A15":"Aménagements, agencements et installations",							
    "A16":"Matériel, mobilier et actifs biologiques",							
    "A17":"Matériel de transport",							
    "A18":"SOUS TOTAL : IMMOBILISATIONS CORPORELLES",							
    "A19":"TOTAL GENERAL",							
    "A20":"[1] I : crédit-bail immobilier; M : crédit-bail mobilier; A : autres contrats (dédoubler le poste si montants significatifs)",						
    "A21":"Commentaire :"	,					
    "A22":"· Indiquer la nature du bien, le nom du bailleur et la durée du bail"
  }
 ];



 for(let i = 1 ; i <= 22; i++){
  NB3alph.forEach((a) => {
    if(`${a}${i}` != "A8"){
      worksheetNB3 .getCell(`${a}${i}`).border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }
    }
    

  })


}

     for(let i = 1 ; i <= 39; i++){
      NB3alph.forEach((a) => {
          if(`${a}${i}` != "A8"){
            worksheetN3.getCell(`${a}${i}`).border =  {
              bottom: {style:'hair', color: {argb:'000000'}},
              left: {style:'hair', color: {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              top: {style:'hair', color: {argb:'000000'}}
            }
          }})}
      for(let i = 0; i <= NB3Data.length; i++ ){
        NB3Data.forEach((v,k) =>{
            for(var d in v)  {
             console.log(d)
              let cell = worksheetNB3.getCell(d);
              cell.value = v[d];
              cell.font = {
                name: 'Cambria',
                size: 9,
                //underline: 'single',
               bold: d  === "A1"? true : false,
                color: { argb: '000000' }
              }
              cell.alignment = { vertical: 'middle', horizontal: d === "A1"? 'center' : 'justify'}
            } })
    
      }





      ///Note 3C  

      const worksheetNC3 = workbook.addWorksheet('Note 3C',{views: [{showGridLines: true}]});
      worksheetNC3 .mergeCells('A1','E2');
      worksheetNC3 .columns = 
      [{width: 40},
       {width: 15},
       {width: 15},
       {width: 15},
       {width: 15}];
      
       const NC3alph =  [
        'A',
        'B',
        'C',
        'D',
        'E',
        ];
      
       const NC3Data = [
         {
          "lignes": 45
         },
         {

"A1":"NOTE 3C : IMMOBILISATIONS (AMORTISSEMENTS)",						
"A3":"Désignation entité :",			
"D3":"Excercice clos le :	",
"E3":"31/12/2019",
"A4":"0"	,		
"D4":"Durée en mois :",
"E4":"12",
"A5":"Numéro d'identification :",
"B5":"004942901",	
"A7":"SITUATIONS ET MOUVEMENTS",	
"B7":"A",
"C7":"B",
"D7":"C",
"E7":"D = A +B - C",
"B8":"Amortissements cumulés à l'ouverture de l'exercice",	
"C8":"Augmentations : Dotations de l'exercice	",
"D8":"Diminutions : Amortissements relatifs aux éléments sortis de l'actif",
"E8":"Cumul des amortissements à la clôture de l'exercice",
"A9":"Frais de développement et de prospection",			
"A10":"Brevets, licences, logiciels et droits similaires",			
"A11":"Fonds commercial et droit au bail",			
"A12":"Autres immobilisations incorporelles",			
"A13":"SOUS TOTAL : IMMOBILISATIONS INCORPORELLES",				
"A14":"Terrains hors immeuble de placement",				
"A15":"Terrains immeuble de placement",				
"A16":"Bâtiments hors immeuble de placement",		
"A17":"Bâtiments immeuble de placement	",			
"A18":"Aménagements, agencements et installations",				
"A19":"Matériel, mobilier et actifs biologiques",			
"A20":"Matériel de transport",				
"A21":"SOUS TOTAL : IMMOBILISATIONS CORPORELLES",				
"A22":"TOTAL GENERAL",			
"A23":"Commentaire :",			
"A24":"· Indiquer :",			
"A25":"- les modes d'amortissements utilisés : amortissement linéaire",				
"A26":"- la durée de vie ou les taux d'amortissements utilisés ;"			

        }
       ];
      
      
      
       for(let i = 1 ; i <= 26; i++){
             
      
        NC3alph.forEach((a) => {
          if(`${a}${i}` != "A8"){
            worksheetNC3 .getCell(`${a}${i}`).border =  {
              bottom: {style:'hair', color: {argb:'000000'}},
              left: {style:'hair', color: {argb:'000000'}},
              right: {style:'hair', color: {argb:'000000'}},
              top: {style:'hair', color: {argb:'000000'}}
            }
          }
          
      
        })
      
      
      }
      
      for(let i = 0; i <= Object.keys(NC3Data).length; i++ ){
        for(var v in NC3Data[1]){
             console.log(`mon ${v}`)
              let cell = worksheetNC3.getCell(v);
             cell.value = NC3Data[1][v];
              cell.font = {
                name: 'Cambria',
                size: 9,
                //underline: 'single',
               // bold: true,
                color: { argb: '000000' }
              }
              cell.alignment = { vertical: 'middle', horizontal: 'justify'}
        }
      
      }




///Note C3B  

const worksheetNC3B = workbook.addWorksheet('Note 3C Bis',{views: [{showGridLines: true}]});
worksheetNC3B .mergeCells('A1','E2');
worksheetNC3B .columns = 
[{width: 40},
 {width: 15},
 {width: 15},
 {width: 15},
 {width: 15}];

 const NC3Balph =  [
  'A',
  'B',
  'C',
  'D',
  'E',
  ];

 const NC3BData = [
   {
    "lignes": 45
   },
   {

"A1":"NOTE 3C : IMMOBILISATIONS (DEPRECIATIONS)",			
"A3":"Désignation entité :",		
"D3":"Excercice clos le :",	
"D4":"Durée en mois :",	
"A5":"Numéro d'identification :",			
"A7":"SITUATIONS ET MOUVEMENTS",	
"B7":"A",	
"C7":"B",	
"D7":"C",	
"E7":"D = A +B - C",
"B8":"Amortissements cumulés à l'ouverture de l'exercice",	
"C8":"Augmentations : Dotations de l'exercice",	
"D8":"Diminutions : Amortissements relatifs aux éléments sortis de l'actif",	
"E8":"Cumul des amortissements à la clôture de l'exercice",
"A10":"Frais de développement et de prospection",				
"A11":"Brevets, licences, logiciels et droits similaires",			
"A12":"Fonds commercial et droit au bail",				
"A13":"Autres immobilisations incorporelles",			
"A14":"SOUS TOTAL : IMMOBILISATIONS INCORPORELLES",				
"A15":"Terrains hors immeuble de placement",				
"A16":"Terrains immeuble de placement",				
"A17":"Bâtiments hors immeuble de placement",				
"A18":"Bâtiments immeuble de placement",			
"A19":"Aménagements, agencements et installations",				
"A20":"Matériel, mobilier et actifs biologiques",			
"A21":"Matériel de transport",				
"A22":"SOUS TOTAL : IMMOBILISATIONS CORPORELLES",				
"A23":"Avances et accomptes versé sur immobilisations incorporelles",				
"A24":"Avances et accomptes versé sur immobilisations corporelles",				
"A25":"SOUS TOTAL : AVANCES ET ACCOMPTES  VERSES",			
"A26":"Titres de participations",				
"A27":"Autres immobilisations financiéres",				
"A28":"SOUS TOTAL : IMMOBILISATIONS FINANCIERES",				
"A30":"TOTAL GENERAL",				
"A31":"Commentaire :",				
"A32":"· Indiquer :",				
"A33":"- les modes d'amortissements utilisés : amortissement linéaire",				
"A34":"- la durée de vie ou les taux d'amortissements utilisés ;"			

  }
 ];



 for(let i = 1 ; i <= 30; i++){
       

  NC3Balph.forEach((a) => {
    if(`${a}${i}` != "A8"){
      worksheetNC3B .getCell(`${a}${i}`).border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }
    }
    

  })


}

for(let i = 0; i <= Object.keys(NC3BData).length; i++ ){

  for(var v in NC3BData[1]){

  
       console.log(`mon ${v}`)
        let cell = worksheetNC3B.getCell(v);
       cell.value = NC3BData[1][v];
        cell.font = {
          name: 'Cambria',
          size: 9,
          //underline: 'single',
         // bold: true,
          color: { argb: '000000' }
        }
  
        cell.alignment = { vertical: 'middle', horizontal: 'justify'}

      
      

  }

}





///Note C3D  

const worksheetNC3D = workbook.addWorksheet('Note 3D',{views: [{showGridLines: true}]});
worksheetNC3D.mergeCells('A1','E2');
worksheetNC3D.columns = 
[{width: 40},
 {width: 20},
 {width: 20},
 {width: 20},
 {width: 20},
 {width: 20}];

 const NC3Dalph =  [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F'
  ];

 const NC3DData = [
   {
    "lignes": 45
   },
   {
"A1":"NOTE 3D : IMMOBILISATIONS (PLUS-VALUES ET MOINS VALUES DE CESSION)",					
"A3":"Désignation entité :",			
"D3":"Excercice clos le :",	
"E3":"31/12/2019",	
"A4":"0",			
"D4":"Durée en mois :",	
"E4":"12",	
"A5":"Numéro d'identification :",	
"B5":"004942901",				
"E5":"MONTANT BRUT",	
"B7":"AMORTISSEMENTS",	
"C7":"VALEUR COMPTABLE NETTE",	
"D7":"PRIX DE CESSION",	
"E7":"PLUS VALUE  \n OU MOINS VALUE",
"C8":"PRATIQUES",		
"B9":"A",	
"C9":"B",	
"D9":"C = A - B",	
"E9":"D",	
"F9":"E = D - C",
"A10":"Frais de développement et de prospection",					
"A11":"Brevets, licences, logiciels et droits similaires",					
"A12":"Fonds commercial et droit au bail",					
"A13":"Autres immobilisations incorporelles",					
"A14":"SOUS TOTAL : IMMOBILISATIONS INCORPORELLES",					
"A15":"Terrains",					
"A16":"Bâtiments",					
"A17":"Aménagements, agencements et installations",					
"A18":"Matériel, mobilier et actifs biologiques",					
"A19":"Matériel de transport",					
"A20":"SOUS TOTAL : IMMOBILISATIONS CORPORELLES",					
"A21":"Titres de placement",					
"A22":"Autres immobilisations financières",					
"A23":"SOUS TOTAL : IMMOBILISATIONS FINANCIERES",					
"A24":"TOTAL GENERAL",					
"A25":"Commentaire :",					
"A26":"· Mentionner la justification de la cession ainsi que la date d'acquisition et la date de sortie",					



  }
 ];



 for(let i = 1 ; i <= 30; i++){
       

  NC3Dalph.forEach((a) => {
    if(`${a}${i}` != "A8"){
      worksheetNC3D .getCell(`${a}${i}`).border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }
    }
    

  })


}

for(let i = 0; i <= Object.keys(NC3DData).length; i++ ){

  for(var v in NC3DData[1]){

  
       console.log(`mon ${v}`)
        let cell = worksheetNC3D.getCell(v);
       cell.value = NC3DData[1][v];
        cell.font = {
          name: 'Cambria',
          size: 9,
          //underline: 'single',
         // bold: true,
          color: { argb: '000000' }
        }
  
        cell.alignment = { vertical: 'middle', horizontal: 'justify'}

      
      

  }

}




///Note C3D  

const worksheetN3E = workbook.addWorksheet('Note 3E',{views: [{showGridLines: true}]});
worksheetN3E.mergeCells('A1','E2');
worksheetN3E.columns = 
[{width: 40},
 {width: 40},
 {width: 40},
 {width: 10}];

 const N3Ealph =  [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F'
  ];

 const N3EData = [
   {
    "lignes": 45
   },
   {
"A1":"NOTE 3E : INFORMATIONS SUR LES REEVALUATIONS EFFECTUEES PAR L'ENTITE",			
"A3":"Désignation entité :",		
"C3":"Excercice clos le :",	
"D3":"31/12/2019",
"A4":"0",		
"C4":"Durée en mois :",	
"D4":"12",
"A5":"Numéro d'identification :",	
"B5":"004942901",		
"A7":"Nature et date des réévaluations :",			
"A13":"Eléments réévalués par postes du bilan",	
"B13":"Montants coûts historiques",	
"C13":"Amortissements supplémentaires",	
"A22":"Méthode de réévaluation utilisée :",			
"A24":"Traitement fiscal de l'écart de réévaluation",			
"A25":"et des amortissements supplémentaires :",		
"A27":"Montant de l'écart incorporé au capital :"			



  }
 ];



 for(let i = 1 ; i <= 27; i++){
       

  N3Ealph.forEach((a) => {
    if(`${a}${i}` != "A8"){
      worksheetN3E .getCell(`${a}${i}`).border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }
    }
    

  })


}

for(let i = 0; i <= Object.keys(N3EData).length; i++ ){

  for(var v in N3EData[1]){

  
       console.log(`mon ${v}`)
        let cell = worksheetN3E.getCell(v);
       cell.value = N3EData[1][v];
        cell.font = {
          name: 'Cambria',
          size: 9,
          //underline: 'single',
         // bold: true,
          color: { argb: '000000' }
        }
  
        cell.alignment = { vertical: 'middle', horizontal: 'justify'}

      
      

  }

}



///Note 4

const worksheetN4 = workbook.addWorksheet('Note 4',{views: [{showGridLines: true}]});
worksheetN4.mergeCells('A1','H2');
worksheetN4.columns = 
[{width: 30},
 {width: 20},
 {width: 20},
 {width: 20},
 {width: 20},
 {width: 15},
 {width: 15},
 {width: 15}];

 const N4alph =  [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H'];

 const N4Data = [
   {
    "lignes": 45
   },
   {
    "A1":"NOTE 4 : IMMOBILISATIONS FINANCIERES",							
    "A3":"Désignation entité :",					
    "F3":"Excercice clos le :",	
    "G3":"31/12/2019",	
    "F4":"Durée en mois :",	
    "G4":"12",	
    "A5":"Numéro d'identification :",
    "B5":"004942901",						
    "A7":"SITUATIONS ET MOUVEMENTS",	
    "B7":"ANNEE 2019",	
    "C7":"ANNEE 2018",	
    "D7":"Variation en valeur absolue",	
    "E7":"Variation en %",	
    "F7":"Créances à un an au plus",	
    "G7":"Créances à plus d'un an et à deux ans au plus",	
    "H7":"Créances à plus de deux ans",
    "A9":"Titres de participation",							
    "A10":"Prêts et créances",							
    "A11":"Prêt au personnel",							
    "A12":"Créances sur l'Etat",						
    "A13":"Créances sur le concédant",						
    "A14":"Titres immobilisés",							
    "A15":"Dépôts et cautionnements",							
    "A16":"Intérêts courus",							
    "A17":"Créances rattachées à des avances et participations à des GIE",							
    "A18":"Immobilisations financières diverse",							
    "A19":"TOTAL BRUT",							
    "A20":"Dépréciations titres de participation",							
    "A21":"Dépréciations autres immobilisations",							
    "A22":"TOTAL NET DE DEPRECIATION",						
    "A24":"Liste des filiales et participations :",							
    "A26":"Dénomination sociale",	
    "B26":"Localisation (ville/pays)",
    "C26":"Valeur d'acquisition",			
    "E26":"% détenu",	
    "G26":"Montant des \n capitaux propres \n filiale",	
    "H26":"Résultat dernier exercice filiale",
    "A35":"Commentaire :",						
    "A36":"· Justifier toute variation significative",							
    "A37":"· Commenter toutes les créances anciennes",						
    "A38":"· Pour les créances relatives à la concession, faire un descriptif de l'accord",							
    "A39":"· Indiquer :",							
    "A40":"- la nature de la créance ;",							
    "A41":"- la durée de la concession ;",						
    "A42":"- l'échéance ;",							
    "A43":"· Indiquer le nombre et la date d'acquisition des actions ou parts propres",							
    "A44":"· Dépréciation : indiquer les évènements et les circonstances qui ont motivé",							
    "A45":"la dépréciation ou la reprise"							
    
  }
 ];



 for(let i = 1 ; i <= 34; i++){
  N4alph.forEach((a) => {
    if(`${a}${i}` != "A8"){
      worksheetN4 .getCell(`${a}${i}`).border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }
    }
  })
}

for(let i = 0; i <= Object.keys(N4Data).length; i++ ){
  for(var v in N4Data[1]){
        let cell = worksheetN4.getCell(v);
       cell.value = N4Data[1][v];
        cell.font = {
          name: 'Cambria',
          size: 9,
          //underline: 'single',
         // bold: true,
          color: { argb: '000000' }
        }
  
        cell.alignment = { vertical: 'middle', horizontal: 'justify'}

      
      

  }

}




///Note 5

const worksheetN5 = workbook.addWorksheet('Note 5',{views: [{showGridLines: true}]});
worksheetN4.mergeCells('A1','H2');
worksheetN4.columns = 
[{width: 30},
 {width: 20},
 {width: 20},
 {width: 20},
 {width: 20},
 {width: 15},
 {width: 15},
 {width: 15}];

 const N5alph =  [
  'A',
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H'];

 const N5Data = [
   {
    "lignes": 45
   },
   {
    "A1":"NOTE 4 : IMMOBILISATIONS FINANCIERES",							
    "A3":"Désignation entité :",					
    "F3":"Excercice clos le :",	
    "G3":"31/12/2019",	
    "F4":"Durée en mois :",	
    "G4":"12",	
    "A5":"Numéro d'identification :",
    "B5":"004942901",						
    "A7":"SITUATIONS ET MOUVEMENTS",	
    "B7":"ANNEE 2019",	
    "C7":"ANNEE 2018",	
    "D7":"Variation en valeur absolue",	
    "E7":"Variation en %",	
    "F7":"Créances à un an au plus",	
    "G7":"Créances à plus d'un an et à deux ans au plus",	
    "H7":"Créances à plus de deux ans",
    "A9":"Titres de participation",							
    "A10":"Prêts et créances",							
    "A11":"Prêt au personnel",							
    "A12":"Créances sur l'Etat",						
    "A13":"Créances sur le concédant",						
    "A14":"Titres immobilisés",							
    "A15":"Dépôts et cautionnements",							
    "A16":"Intérêts courus",							
    "A17":"Créances rattachées à des avances et participations à des GIE",							
    "A18":"Immobilisations financières diverse",							
    "A19":"TOTAL BRUT",							
    "A20":"Dépréciations titres de participation",							
    "A21":"Dépréciations autres immobilisations",							
    "A22":"TOTAL NET DE DEPRECIATION",						
    "A24":"Liste des filiales et participations :",							
    "A26":"Dénomination sociale",	
    "B26":"Localisation (ville/pays)",
    "C26":"Valeur d'acquisition",			
    "E26":"% détenu",	
    "G26":"Montant des \n capitaux propres \n filiale",	
    "H26":"Résultat dernier exercice filiale",
    "A35":"Commentaire :",						
    "A36":"· Justifier toute variation significative",							
    "A37":"· Commenter toutes les créances anciennes",						
    "A38":"· Pour les créances relatives à la concession, faire un descriptif de l'accord",							
    "A39":"· Indiquer :",							
    "A40":"- la nature de la créance ;",							
    "A41":"- la durée de la concession ;",						
    "A42":"- l'échéance ;",							
    "A43":"· Indiquer le nombre et la date d'acquisition des actions ou parts propres",							
    "A44":"· Dépréciation : indiquer les évènements et les circonstances qui ont motivé",							
    "A45":"la dépréciation ou la reprise"							
    
  }
 ];



 for(let i = 1 ; i <= 34; i++){
  N5alph.forEach((a) => {
    if(`${a}${i}` != "A8"){
      worksheetN5 .getCell(`${a}${i}`).border =  {
        bottom: {style:'hair', color: {argb:'000000'}},
        left: {style:'hair', color: {argb:'000000'}},
        right: {style:'hair', color: {argb:'000000'}},
        top: {style:'hair', color: {argb:'000000'}}
      }
    }
  })
}

for(let i = 0; i <= Object.keys(N5Data).length; i++ ){
  for(var v in N5Data[1]){
        let cell = worksheetN5.getCell(v);
       cell.value = N4Data[1][v];
        cell.font = {
          name: 'Cambria',
          size: 9,
          //underline: 'single',
         // bold: true,
          color: { argb: '000000' }
        }
  
        cell.alignment = { vertical: 'middle', horizontal: 'justify'}

      
      

  }

}

       


  
    


        








          







    









    






    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob,  'exported.xlsx');
    })

    

  }
}
