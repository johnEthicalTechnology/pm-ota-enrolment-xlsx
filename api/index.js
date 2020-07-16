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
  console.log('1) Zoho object parsed into JS object');

  // 39 fields from Zoho Course Record object
  const facilitator = course_record.New_Facilitator_Name
  const additionalNotes = course_record.Other_Notes
  const anyFurtherNotes = course_record.Data_Projection_Availability_Notes
  const organisation = course_record.Company.name
  const keyDecisionMaker = course_record.KDM_Name
  const pmAccountManager = course_record.Owner.name
  const invoiceContactName = course_record.IC_Name
  const invoicePhone = course_record.IC_Phone
  const invoiceEmail = course_record.IC_Email
  const trainingProgram = course_record.Product_Name
  const version = course_record.Course_Version
  const coaching = course_record.Coaching_Included
  const deliveryMode = course_record.Course_Delivery
  const trainingDates = parseTrainingDates(course_record)
  const startTime = course_record.Start_Time
  const finishTime = course_record.Finish_Time
  // TODO - Ask Mario - Need to add trainingVenueAndRoom fields to Zoho OTA
  // const trainingVenueAndRoom = course_record
  const address = `${course_record.VD_Street_Address}, ${course_record.VD_Suburb}, ${course_record.VD_State}, ${course_record.VD_Postcode}`
  const venueContactPerson = course_record.VD_Name
  const venueMobileNo = course_record.VD_Phone
  const venueOfficeNo = course_record.VD_Office_Phone
  const pmEarlyAccess = course_record.PM_Early_Access
  const trainerAccessTime = course_record.Facilitator_Access_Time
  const parkingAvailable = course_record.Parking_Available
  const siteInductionOrPpe = course_record.Site_Induction_or_PPE_Required
  const nameAndAddressReceiveTrainingMaterials = `${course_record.MD_Name}, ${course_record.MD_Street_Address}, ${course_record.MD_Suburb}, ${course_record.MD_State}, ${course_record.MD_Postcode}`
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
  const cateringProvided = course_record.Catering_Provided
  const morningTea = course_record.Morning_Tea_Time
  const lunch = course_record.Lunch_Time
  const afternooTea = course_record.Afternoon_Tea_Time
  const trainingMaterialsProvided = course_record.Training_Resources_Provided_by_Priority_Management
  // Read templated sheet
  const facilitatorWb = new ExcelJS.Workbook()
  const facilitatorWs = await facilitatorWb.xlsx.readFile(join(__dirname, '_files', 'otaAndEnrolmentSheet.xlsx'))
  // Map Zoho fields to Spreadsheet
  const otaSheet = facilitatorWs.getWorksheet('OTA')
  otaSheet.getCell('E4').value = facilitator
  otaSheet.getCell('E5').value = additionalNotes
  otaSheet.getCell('D6').value = anyFurtherNotes
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
  otaSheet.getCell('A19').value = course_record.Course_Type == 'Custom' ? courseCustomisation : ''
  // otaSheet.getCell('A22').value = trainingVenueAndRoom
  otaSheet.getCell('E24').value = address
  otaSheet.getCell('A24').value = venueContactPerson
  otaSheet.getCell('E24').value = venueMobileNo
  otaSheet.getCell('G24').value = venueOfficeNo
  otaSheet.getCell('H25').value = trainerAccessTime
  otaSheet.getCell('A27').value = parkingAvailable == true ? course_record.Special_Conditions_for_Parking : 'No'
  otaSheet.getCell('E27').value = siteInductionOrPpe == true ? 'Yes' : 'No'
  otaSheet.getCell('A29').value = nameAndAddressReceiveTrainingMaterials
  otaSheet.getCell('A32').value = course_record.Mailbox_Limit ? mailboxLimit : 'No'
  otaSheet.getCell('E32').value = course_record.Existing_Archive_Policy ?existingArchivePolicy : 'No'
  otaSheet.getCell('E34').value = course_record.CRM_or_Workload_Management_System == true ? crmOrWmSystem : 'No'
  otaSheet.getCell('A34').value = mobileDevicesInUse != null ? mobileDevicesInUse : 'No'
  otaSheet.getCell('A36').value = departmentBeingTrained
  otaSheet.getCell('A41').value = computerSupplier
  otaSheet.getCell('I41').value = pmEarlyAccess
  otaSheet.getCell('C42').value = dataProjector == true ? 'Yes' : 'No'
  otaSheet.getCell('E42').value = screen == true ? 'Yes' : 'No'
  otaSheet.getCell('H42').value = whiteboard == true ? 'Yes' : 'No'
  otaSheet.getCell('J42').value = flipchart == true ? 'Yes' : 'No'
  otaSheet.getCell('J43').value = cateringProvided == true ? 'Yes' : 'No'
  otaSheet.getCell('B44').value = morningTea
  otaSheet.getCell('F44').value = lunch
  otaSheet.getCell('I44').value = afternooTea
  otaSheet.getCell('F45').value = trainingMaterialsProvided.join(', ')
  console.log('2) Added Zoho Course OTA to spreadsheet');

  const attendanceSheet = facilitatorWs.getWorksheet('attendance')
  const START_OF_ATTENDANTS_LIST = 4
  attendanceSheet.getCell('C1').value = trainingProgram
  const startDateOfCourse = toDateString(course_record.Date_of_Training)
  attendanceSheet.getCell('C2').value = startDateOfCourse
  attendee_map_list.forEach((attendeeDetails, index) => {
    console.log(`3)a) Adding ${attendeeDetails}`);
    attendanceSheet.getCell(`A${index + START_OF_ATTENDANTS_LIST}`).value = index + 1
    attendanceSheet.getCell(`B${index + START_OF_ATTENDANTS_LIST}`).value = attendeeDetails.firstName
    attendanceSheet.getCell(`C${index + START_OF_ATTENDANTS_LIST}`).value = attendeeDetails.lastName
    attendanceSheet.getCell(`D${index + START_OF_ATTENDANTS_LIST}`).value = attendeeDetails.jobTitle
    attendanceSheet.getCell(`E${index + START_OF_ATTENDANTS_LIST}`).value = attendeeDetails.emailAddress
  })
  console.log('3)b) Added attendee details to attendanceSheet');

  try {
    //* 4) Create buffer
    const buffer = await facilitatorWb.xlsx.writeBuffer()
    //* 5) Create reusable transporter object using the default SMTP transport
    const transporter = nodemailer.createTransport({
      host: 'smtp.zoho.com',
      port: 465,
      secure: true, // true for 465, false for other ports
      auth: {
        user: 'brett.handley@prioritymanagement.com.au',
        pass: process.env.EMAIL_PW,
      },
    })

    //* 6) Send mail with defined transport object
    const emailRes = await transporter.sendMail({
      from: `'Priority Management Sydney' <brett.handley@prioritymanagement.com.au>`,
      to: 'materials@prioritymanagement.com.au',
      subject: `Spreadsheet for facilitator - ${facilitator}`,
      text: `Dear PM Admin,/r This is the Facilitator spreadsheet for the course ${trainingProgram} held by ${facilitator} and starting on ${startDateOfCourse}/r Regards,`,
      html: `<p>Dear PM Admin,</p><p>This is the Facilitator spreadsheet for the course ${trainingProgram} held by ${facilitator} and starting on ${startDateOfCourse}</p><p>Regards,</p>`,
      attachments: [
        {
          filename: `${course_record.Date_of_Training}-${trainingProgram}.xlsx`,
          content: buffer
        }
      ]
    })
    console.log('Message sent:', emailRes.messageId)
    res.json({body: `Message sent: ${emailRes.messageId}`})
  } catch (error) {
    console.error('Error:', error)
    res.json({body: `Error: ${error}`})
  }
}