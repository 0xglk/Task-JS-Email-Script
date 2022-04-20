const moment = require('moment');
const nodemailer = require('nodemailer');
let date_now = moment().format('DD/MM/YYYY'); 
// Your code goes here
'use strict';
const excelToJson = require('convert-excel-to-json');
 
const result = excelToJson({
    sourceFile: 'names.xlsx',
    columnToKey: {
        A: 'Name',
        B: 'Email',
        C: 'Course',
        D: 'Grade'
    },
    header:{
        rows: 1
    }
});
 
const transporter = nodemailer.createTransport({
    host: 'smtp.gmail.com',
    port: 587,
    auth: {
      user: 'Your E-mail',
      pass: 'Your Password',
    },
  });

  transporter.verify().then(console.log).catch(console.error);

const information = (result[Object.keys(result)[0]]);
information.forEach(element => {
    transporter.sendMail({
        from: '"Your Name" <Email@gmail.com>',
        to: `${element["Email"]}`,
        subject: 'Congratulations!🎉',
        html: `
        <div style="width:90%; height:90%; padding:20px; text-align:center; border: 10px solid #787878">
        <div style="width:90%; height:90%; padding:20px; text-align:center; border: 5px solid #787878">
        <span style="font-size:40px; font-weight:bold">Certificate of Completion</span>
        <br><br>
        <span style="font-size:25px"><i>This is to certify that</i></span>
        <br><br>
        <span style="font-size:30px"><b>${element["Name"]}</b></span><br/><br/>
        <span style="font-size:25px"><i>has completed the course</i></span> <br/><br/>
        <span style="font-size:30px">${element["Course"]}</span> <br/><br/>
        <span style="font-size:20px">with score of <b>${element["Grade"]}</b></span> <br/><br/><br/><br/>
        <span style="font-size:25px"><i>dated</i></span><br>
        <span style="font-size:15px">${date_now}</span>
        </div>
        </div> 
        `,
      }).then(info => {
        console.log({info});
      }).catch(console.error);
});
