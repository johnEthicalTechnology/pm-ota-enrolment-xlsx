const nodemailer = require('nodemailer')
const ExcelJS = require('exceljs')
const { join } = require('path')

function toDateString(dayOfTrainingDate) {
  const dateSplit = dayOfTrainingDate.split('-')
  const dateString = new Date(dateSplit[0], dateSplit[1], dateSplit[2]).toDateString()
  return dateString
}

function parseTrainingDates(courseRecord) {
  const daysOfTraining = courseRecord.Days_of_Training
  let trainingDates
  if(daysOfTraining >= 1) {
    trainingDates = toDateString(courseRecord.Date_of_Training)
  }
  if(daysOfTraining >= 2) {
    trainingDates += `, ${toDateString(courseRecord.Day_Two_Date)}`
  }
  if(daysOfTraining >= 3) {
    trainingDates += `, ${toDateString(courseRecord.Day_Three_Date)}`
  }
  if(daysOfTraining >= 4) {
    trainingDates += `, ${toDateString(courseRecord.Day_Four_Date)}`
  }
  if(daysOfTraining >= 5) {
    trainingDates += `, ${toDateString(courseRecord.Day_Five_Date)}`
  }
  return trainingDates
}

module.exports = async (req, res) => {
  const { course_record, attendee_map_list } = JSON.parse(req.body.data)
  // 39 fields
  const facilitator = course_record.Facilitator_Name
  const additionalNotes
  // TODO - ask about this unsure about this one
  // const furtherNotes
  const organisation = course_record.Company.name
  const keyDecisionMaker = KDM_Name
  const pmAccountManager = Owner.name
  const invoiceContactName = IC_Name
  const invoicePhone = IC_Phone
  const invoiceEmail = IC_Email
  const trainingProgram = Product_Name
  const version = Course_Version
  const coaching = Coaching_Included
  const deliveryMode = Course_Delivery
  const trainingDates = parseTrainingDates(course_record)
  const startTime = course_record.Start_Time
  const finishTime = course_record.Finish_Time
  // TODO - Ask Mario - Need to add trainingVenueAndRoom fields to Zoho OTA
  // const trainingVenueAndRoom = course_record
  const address = `${course_record.VD_Street_Address}, ${course_record.VD_Suburb}, ${course_record.VD_State}, ${course_record.VD_Postcode}`
  const venueContactPerson = course_record.VD_Name
  const venueMobileNo = course_record.VD_Phone
  // TODO - Add VD_Office_Number
  const venueOfficeNo = course_record.VD_Office_Number
  const pmEarlyAccess = course_record.PM_Early_Access
  const trainerAccessTime = course_record.Facilitator_Access_Time
  const parkingAvailable = course_record.Parking_Available
  const siteInductionOrPpe = course_record.Site_Induction_or_PPE_Required
  const nameAndAddressReceiveTrainingMaterials = `${course_record.MD_Name}, ${course_record.MD_Address}, ${course_record.MD_Suburb}, ${course_record.MD_State}, ${course_record.MD_Postcode}`
  const mailboxLimit = course_record.Note_for_Mailbox_Limit
  const existingArchivePolicy = course_record.Note_for_Existing_Archive_Policy
  const crmOrWmSystem = course_record.Note_for_CRM_or_Workload_Management_System
  const mobileDevicesInUse = course_record.Mobile_Devices_in_Use
  const departmentBeingTrained = course_record.Department_Being_Trained
  const courseCustomisation = course_record.Type_of_Course_Customisation
  const computerSupplier = course_record.Computer_Supplier
  const dataProjector = course_record.Data_Projector_Available
  const screen = course_record.Laptop_Connectable_Monitor_or_TV_Available
  const whiteboard = course_record.Whiteboard_Available
  const flipchart = course_record.Flipchart_Available
  const morningTea = course_record.Morning_Tea_Time
  const lunch = course_record.Lunch_Time
  const afternooTea = course_record.Afternoon_Tea_Time
  const trainingMaterialsProvided = course_record.Training_Resources_Provided_by_Priority_Management

  const otaSheet = worksheet.getWorksheet('OTA')
  otaSheet.getCell('E4').value = facilitator
  otaSheet.getCell('E5').value = additionalNotes
  // otaSheet.getCell('D6').value = furtherNotes
  otaSheet.getCell('A9').value = organisation
  otaSheet.getCell('E9').value = keyDecisionMaker
  otaSheet.getCell('E10').value = pmAccountManager
  otaSheet.getCell('E11').value = invoiceContactName
  otaSheet.getCell('I11').value = invoicePhone
  otaSheet.getCell('F12').value = invoiceEmail
  otaSheet.getCell('A14').value = trainingProgram
  otaSheet.getCell('E14').value = version
  otaSheet.getCell('G14').value = coaching
  otaSheet.getCell('I14').value = deliveryMode
  otaSheet.getCell('A17').value = trainingDates
  otaSheet.getCell('F17').value = startTime
  otaSheet.getCell('I17').value = finishTime
  otaSheet.getCell('A20').value = trainingVenueAndRoom
  otaSheet.getCell('E20').value = address
  otaSheet.getCell('A22').value = venueContactPerson
  otaSheet.getCell('E22').value = venueMobileNo
  otaSheet.getCell('G22').value = venueOfficeNo
  otaSheet.getCell('H23').value = trainerAccessTime
  otaSheet.getCell('A25').value = parkingAvailable
  otaSheet.getCell('E25').value = siteInductionOrPpe
  otaSheet.getCell('A27').value = nameAndAddressReceiveTrainingMaterials
  otaSheet.getCell('A30').value = course_record.Mailbox_Limit ? mailboxLimit : ''
  otaSheet.getCell('E30').value = course_record.Existing_Archive_Policy ?existingArchivePolicy : ''
  otaSheet.getCell('E32').value = course_record.CRM_or_Workload_Management_System ? crmOrWmSystem : ''
  otaSheet.getCell('A32').value = mobileDevicesInUse == null ? 'No' : 'Yes'
  otaSheet.getCell('A34').value = departmentBeingTrained
  otaSheet.getCell('A36').value = course_record.Course_Type == 'Custom' ? courseCustomisation : ''
  otaSheet.getCell('A39').value = computerSupplier
  otaSheet.getCell('I39').value = pmEarlyAccess
  otaSheet.getCell('C40').value = dataProjector
  otaSheet.getCell('E40').value = screen
  otaSheet.getCell('H40').value = whiteboard
  otaSheet.getCell('J40').value = flipchart
  otaSheet.getCell('B42').value = morningTea
  otaSheet.getCell('F42').value = lunch
  otaSheet.getCell('I42').value = afternooTea
  otaSheet.getCell('F43').value = trainingMaterialsProvided

  test.xlsx.writeFile('testing.xlsx')
  const workbook = new ExcelJS.Workbook()
  const worksheet = await workbook.xlsx.readFile(join(__dirname, '_files', 'otaAndEnrolmentSheet.xlsx'))

  const attendanceSheet = worksheet.getWorksheet('attendance')
  res.json({body: 'This was a success!'})
}