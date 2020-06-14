const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');

module.exports = async (req, res) => {
  const workbook = new ExcelJS.Workbook()
  const worksheet = await workbook.xlsx.readFile('otaAndEnrolmentSheet.xlsx')
  const otaSheet = worksheet.getWorksheet('OTA')
  const attendanceSheet = worksheet.getWorksheet('attendance')
  console.log('attendanceSheet', attendanceSheet)
  console.log('otaSheet', otaSheet)
  console.log('req', req.body.data);
  res.json({body: 'This was a success!'})
}