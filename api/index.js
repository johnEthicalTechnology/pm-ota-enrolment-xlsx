const nodemailer = require('nodemailer')
const ExcelJS = require('exceljs')
const { join } = require('path')
const crypto = require('crypto')

function authorise(authorizationToken) {
  const secretKey = process.env.SECRET_KEY
  const secretValue = process.env.SECRET_VALUE
  if (
    authorizationToken ===
    crypto.createHmac('SHA256', secretKey).update(secretValue).digest('base64')
  ) {
    console.log('Authentication successful')
    return true
  } else {
    console.log('Authentication failure')
    return false
  }
}

const dayOfWeekObject = {
  0: 'Sun',
  1: 'Mon',
  2: 'Tue',
  3: 'Wed',
  4: 'Thu',
  5: 'Fri',
  6: 'Sat',
}

function toDateString(dayOfTrainingDate) {
  const dateSplit = dayOfTrainingDate.split('-')
  const dateString = new Date(
    dateSplit[0],
    dateSplit[1] - 1,
    dateSplit[2]
  ).toDateString()
  return dateString
}

function parseTrainingDatesAndTimes(courseRecord) {
  const daysOfTraining = courseRecord.Days_of_Training
  let trainingDates
  let trainingStartTimes
  let trainingFinTimes
  let dayOfWeek
  if (daysOfTraining >= 1) {
    if (daysOfTraining == 1) {
      trainingDates = toDateString(courseRecord.Date_of_Training)
    } else {
      dayOfWeek = new Date(courseRecord.Date_of_Training).getDay()
      trainingDates = `${dayOfWeekObject[dayOfWeek]} ${new Date(
        courseRecord.Date_of_Training
      ).getDate()}`
    }
    trainingStartTimes = `Day 1: ${courseRecord.Start_Time}\r\n`
    trainingFinTimes = `Day 1: ${courseRecord.Finish_Time}\r\n`
  }
  if (daysOfTraining >= 2) {
    if (daysOfTraining == 2) {
      trainingDates += ` & ${toDateString(courseRecord.Day_Two_Date)}`
    } else {
      dayOfWeek = new Date(courseRecord.Day_Two_Date).getDay()
      trainingDates += `, ${dayOfWeekObject[dayOfWeek]} ${new Date(
        courseRecord.Day_Two_Date
      ).getDate()}`
    }
    trainingStartTimes += `Day 2: ${courseRecord.Day_Two_Start_Time}\r\n`
    trainingFinTimes += `Day 2: ${courseRecord.Day_Two_Finish_Time}\r\n`
  }
  if (daysOfTraining >= 3) {
    if (daysOfTraining == 3) {
      trainingDates += ` & ${toDateString(courseRecord.Day_Three_Date)}`
    } else {
      dayOfWeek = new Date(courseRecord.Day_Three_Date).getDay()
      trainingDates += `, ${dayOfWeekObject[dayOfWeek]} ${new Date(
        courseRecord.Day_Three_Date
      ).getDate()}`
    }
    trainingStartTimes += `Day 3: ${courseRecord.Day_Three_Start_Time}\r\n`
    trainingFinTimes += `Day 3: ${courseRecord.Day_Three_Finish_Time}\r\n`
  }
  if (daysOfTraining >= 4) {
    if (daysOfTraining == 4) {
      trainingDates += ` & ${toDateString(courseRecord.Day_Four_Date)}`
    } else {
      dayOfWeek = new Date(courseRecord.Day_Four_Date).getDay()
      trainingDates += `, ${dayOfWeekObject[dayOfWeek]} ${new Date(
        courseRecord.Day_Four_Date
      ).getDate()}`
    }
    trainingStartTimes += `Day 4: ${courseRecord.Day_Four_Start_Time}\r\n`
    trainingFinTimes += `Day 4: ${courseRecord.Day_Four_Finish_Time}\r\n`
  }
  if (daysOfTraining >= 5) {
    if (daysOfTraining == 5) {
      trainingDates += ` & ${toDateString(courseRecord.Day_Four_Date)}`
    }
    trainingStartTimes += `Day 5: ${courseRecord.Day_Five_Start_Time}\r\n`
    trainingFinTimes += `Day 5: ${courseRecord.Day_Five_Finish_Time}\r\n`
  }
  return { trainingDates, trainingStartTimes, trainingFinTimes }
}

module.exports = async (req, res) => {
  const isAuthenticated = authorise(req.headers.authorization, res)
  if (isAuthenticated) {
    const { course_record, attendee_map_list } = JSON.parse(req.body.data)
    console.log('1) Zoho object parsed into JS object')
    console.log('Course record', course_record)
    console.log('Attendee map list', attendee_map_list)

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
    const { trainingDates, trainingStartTimes, trainingFinTimes } =
      parseTrainingDatesAndTimes(course_record)
    // TODO - Ask Mario - Need to add trainingVenueAndRoom fields to Zoho OTA
    // const trainingVenueAndRoom = course_record
    const address = `${course_record.VD_Street_Address}, ${course_record.VD_Suburb}, ${course_record.VD_State}, ${course_record.VD_Postcode}`
    const venueContactPerson = course_record.VD_Name
    const venueMobileNo = course_record.VD_Phone
    const venueOfficeNo = course_record.VD_Office_Phone
    const pmEarlyAccess = course_record.PM_Early_Access
    const trainerAccessTime = course_record.Facilitator_Access_Time
    const parkingAvailable = course_record.Parking_Available_1
    const parkingDetails =
      course_record.Special_Conditions_for_Parking != null
        ? course_record.Special_Conditions_for_Parking
        : ''
    const siteInductionOrPpe = course_record.Site_Induction_or_PPE_Required
    const nameAndAddressReceiveTrainingMaterials = `${course_record.MD_Name}, ${course_record.MD_Street_Address}, ${course_record.MD_Suburb}, ${course_record.MD_State}, ${course_record.MD_Postcode}`
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
    const trainingMaterialsProvided =
      course_record.Training_Resources_Provided_by_Priority_Management
    // Read templated sheet
    const facilitatorWb = new ExcelJS.Workbook()
    const facilitatorWs = await facilitatorWb.xlsx.readFile(
      join(__dirname, '_files', 'otaAndEnrolmentSheet.xlsx')
    )
    // Map Zoho fields to Spreadsheet
    const otaSheet = facilitatorWs.getWorksheet('OTA')
    // Important things to note about this course
    otaSheet.getCell('E4').value = facilitator
    otaSheet.getCell('E5').value = additionalNotes
    otaSheet.getCell('D6').value = anyFurtherNotes
    // WORKSHOP DETAILS
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
    otaSheet.getCell('F17').value = trainingStartTimes
    otaSheet.getCell('I17').value = trainingFinTimes
    otaSheet.getCell('A19').value =
      course_record.Course_Type == 'Custom' ? courseCustomisation : ''
    // VENUE DETAILS
    // otaSheet.getCell('A22').value = trainingVenueAndRoom
    otaSheet.getCell('E22').value = address
    otaSheet.getCell('A24').value = venueContactPerson
    otaSheet.getCell('E24').value = venueMobileNo
    otaSheet.getCell('G24').value = venueOfficeNo
    otaSheet.getCell('H25').value = trainerAccessTime
    otaSheet.getCell('A27').value =
      parkingAvailable == 'Yes' ? `Yes, ${parkingDetails}` : 'No'
    otaSheet.getCell('E27').value = siteInductionOrPpe == 'Yes' ? 'Yes' : 'No'
    otaSheet.getCell('A29').value = nameAndAddressReceiveTrainingMaterials
    // TRAINING SESSION DETAILS
    otaSheet.getCell('A32').value = departmentBeingTrained
    // TRAINING RESOURCES
    otaSheet.getCell('A35').value = computerSupplier
    otaSheet.getCell('I35').value = pmEarlyAccess
    otaSheet.getCell('C36').value = dataProjector == 'Yes' ? 'Yes' : 'No'
    otaSheet.getCell('E36').value = screen == 'Yes' ? 'Yes' : 'No'
    otaSheet.getCell('H36').value = whiteboard == 'Yes' ? 'Yes' : 'No'
    otaSheet.getCell('J36').value = flipchart == 'Yes' ? 'Yes' : 'No'
    otaSheet.getCell('J37').value = cateringProvided == 'Yes' ? 'Yes' : 'No'
    otaSheet.getCell('B38').value = morningTea
    otaSheet.getCell('F38').value = lunch
    otaSheet.getCell('I38').value = afternooTea
    otaSheet.getCell('F39').value = trainingMaterialsProvided.join(', ')
    console.log('2) Added Zoho Course OTA to spreadsheet')

    const attendanceSheet = facilitatorWs.getWorksheet('attendance')
    const START_OF_ATTENDANTS_LIST = 4
    attendanceSheet.getCell('C1').value = trainingProgram
    const startDateOfCourse = toDateString(course_record.Date_of_Training)
    attendanceSheet.getCell('C2').value = startDateOfCourse
    attendee_map_list.forEach((attendeeDetails, index) => {
      console.log(`3)a) Adding ${attendeeDetails.firstName}`)
      attendanceSheet.getCell(`A${index + START_OF_ATTENDANTS_LIST}`).value =
        index + 1
      attendanceSheet.getCell(`B${index + START_OF_ATTENDANTS_LIST}`).value =
        attendeeDetails.firstName
      attendanceSheet.getCell(`C${index + START_OF_ATTENDANTS_LIST}`).value =
        attendeeDetails.lastName
      attendanceSheet.getCell(`D${index + START_OF_ATTENDANTS_LIST}`).value =
        attendeeDetails.jobTitle
      attendanceSheet.getCell(`E${index + START_OF_ATTENDANTS_LIST}`).value =
        attendeeDetails.emailAddress
    })
    console.log('3)b) Added attendee details to attendanceSheet')

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
            content: buffer,
          },
        ],
      })
      console.log('Message sent:', emailRes.messageId)
      res.json({ body: `Message sent: ${emailRes.messageId}` })
    } catch (error) {
      console.error('Error:', error)
      res.json({ body: `Error: ${error}` })
    }
  } else {
    res.json({ body: 'Authentication failure' })
  }
}
