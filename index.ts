
 const moment =require('moment');
const Excel=require("exceljs/dist/exceljs.min.js");


function excels(data:any,excelname:any,img:any,key:any,col:any,title:any,headerNum:any,numLopp:any,imgSize:any,x:any,column_align = []){

 var footer = ["UNIG -[ BWATER]"];

   var workbook = new Excel.Workbook();
   workbook.creator = 'Web';
   workbook.lastModifiedBy = 'Web';
   workbook.created = new Date();
   workbook.modified = new Date();
   workbook.addWorksheet(title, { views: [{ state: 'frozen', ySplit: 5, xSplit: x, activeCell: 'C2', showGridLines: true }] })
   var sheet = workbook.getWorksheet(1);
  /*  var imageId1 = workbook.addImage({
     base64:img,
     extension:'png',
   }); */

   /* if(imgSize>40){
     sheet.addImage(imageId1, {
     tl: { col: 0, row: 1.3 },
     ext: { width: 120, height:80 }
   })
   }else{
     sheet.addImage(imageId1, {
       tl: { col: 0, row: 1.3 },
       ext: { width: 110, height:80 }
     })
   }
 */
sheet.getCell('B3').value = title

   sheet.addRow("");
   sheet.getRow(5).values = col;
   sheet.columns = key;
   sheet.getRow(5).fill = {
     type: 'pattern',
     pattern: 'solid',
     fgColor: { argb: 'ffffff' },
     size: 16
   }

    sheet.addRows(data);


    sheet.addRow('').fill = {
     type: 'pattern',
     pattern: 'solid',
     fgColor: { argb: 'ffffff' },
     size: 26
   };

   sheet.eachRow({ includeEmpty: true }, function (row:any, rowNumber:any) {
     row.eachCell(function (cell:any, colNumber:any) {
       cell.font = {
         name: 'Arial',
         family: 2,
         bold: true,
         size: 26,

       };
       cell.alignment = {
         vertical: 'middle', horizontal: 'center'
       };
       if(rowNumber < headerNum){
         for (var i = 0; i < headerNum; i++) {
         sheet.getRow(i).fill = {
           type: 'pattern',
           pattern: 'solid',
           fgColor: { argb: 'ffffff' },
           size: 26
         }
       }
       }

       if (rowNumber <= headerNum+1) {
         row.height = 20;
         cell.font = {
           bold: true,
           size: 20,
           color: { argb: '0099FF' },
         };
         cell.alignment = {
           vertical: 'middle', horizontal: 'center'
         };
       }

       if (rowNumber >= headerNum ) {


         for (var i = 1; i < numLopp+1; i++) {
           if (rowNumber<headerNum) {
             cell.font = {
               color: { argb: '0099FF' },
               bold: true,
               size:14
             };
             row.height = 25;
             row.getCell(i).fill = {
               type: 'pattern',
               pattern: 'solid',
               fgColor: { argb: 'ffffff' }
             };

             cell.alignment = {
               vertical: 'middle', horizontal: 'center'
             };
             }
           if (rowNumber ==headerNum && rowNumber<headerNum+1) {
             cell.font = {
               color: { argb: 'ffffff' },
               bold: true,
               size:14
             };
             row.height = 25;
             row.getCell(i).fill = {
               type: 'pattern',
               pattern: 'solid',
               fgColor: { argb: '0099FF' }
             };

             cell.alignment = {
               vertical: 'middle', horizontal: 'center'
             };
             }else{
               row.getCell(i).fill = {
               type: 'pattern',
               pattern: 'solid',
               fgColor: { argb: 'ffffff' }
             };

             cell.font = {
               color: { argb: '2e2e2f' },
               bold: false,
               size:12
             };
             cell.alignment = {
               vertical: 'middle', horizontal: 'center'
            };
           }

           row.getCell(i).border = {
             top: { style: 'thin' },
             left: { style: 'thin' },
             bottom: { style: 'thin' },
             right: { style: 'thin' }
           };
         }
       }

       if (rowNumber >= 6) {

           column_align.map(key => {
             row.getCell(key).alignment = {
               vertical: 'middle', horizontal: 'right', 'color': { 'argb': 'FFFF6600' }
             };
         });
          

         }
     });
   });
   
   workbook.xlsx.writeBuffer().then((Data:any) => {
     var blob = new Blob([Data]);

     var url = window.URL.createObjectURL(blob);
     var a = document.createElement("a");
     document.body.appendChild(a);
     a.href = url;
     a.download = excelname;
     a.click();
   });
}