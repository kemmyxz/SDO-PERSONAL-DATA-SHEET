const express = require("express");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;

app.use(bodyParser.json({ limit: "100mb" })); // Increase limit to 100MB
app.use(bodyParser.urlencoded({ extended: true, limit: "100mb" })); // Increase limit to 100MB

app.use(express.static("public"));

const filePath = path.join(__dirname, "PersonalDataSheet.xlsx");

const columns = [
  { header: "SURNAME", key: "surname", width: 20 }, // id="surname"
  { header: "FIRST NAME", key: "firstname", width: 20 }, // id="firstname"
  { header: "MIDDLE NAME", key: "middlename", width: 20 }, // id="middlename"
  { header: "NAME EXTENSION", key: "nameExtension", width: 15 }, // id="nameExtension"
  { header: "DATE OF BIRTH", key: "dateOfBirth", width: 15 }, // id="dateOfBirth"
  { header: "PLACE OF BIRTH", key: "placeOfBirth", width: 20 }, // id="placeOfBirth"
  { header: "SEX", key: "sex", width: 10 }, // id="sexMale", id="sexFemale" (radio group)
  { header: "CIVIL STATUS", key: "civilStatus", width: 15 }, // id="civilSingle", id="civilMarried", id="civilWidowed", id="civilSeparated", id="civilOther" (radio group)
  { header: "CIVIL STATUS OTHER", key: "civilOtherInput", width: 15 }, // id="civilOtherInput"
  { header: "CITIZENSHIP", key: "citizenship", width: 15 }, // id="citizenshipFilipino", id="citizenshipDual" (checkbox group)
  { header: "DUAL TYPE", key: "dualType", width: 15 }, // id="dualByBirth", id="dualByNaturalization" (radio group)
  { header: "DUAL COUNTRY", key: "dualCountry", width: 15 }, // id="dualCountry"
  { header: "HEIGHT", key: "height", width: 10 }, // id="height"
  { header: "WEIGHT", key: "weight", width: 10 }, // id="weight"
  { header: "BLOOD TYPE", key: "bloodType", width: 10 }, // id="bloodType"
  { header: "GSIS ID NO.", key: "gsisId", width: 15 }, // id="gsisId"
  { header: "PAG-IBIG ID NO.", key: "pagibigId", width: 15 }, // id="pagibigId"
  { header: "PHILHEALTH NO.", key: "philhealthNo", width: 15 }, // id="philhealthNo"
  { header: "SSS NO.", key: "sssNo", width: 15 }, // id="sssNo"
  { header: "TIN NO.", key: "tinNo", width: 15 }, // id="tinNo"
  { header: "AGENCY EMPLOYEE NO.", key: "agencyEmpNo", width: 20 }, // id="agencyEmpNo"
  { header: "RESIDENTIAL ADDRESS", key: "residentialAddress", width: 30 }, // id="resAddress1", id="resAddress2", id="resAddress3", id="resZip"
  { header: "PERMANENT ADDRESS", key: "permanentAddress", width: 30 }, // id="permAddress1", id="permAddress2", id="permAddress3", id="permZip"
  { header: "TELEPHONE NO.", key: "telNo", width: 15 }, // id="telNo"
  { header: "MOBILE NO.", key: "mobileNo", width: 15 }, // id="mobileNo"
  { header: "EMAIL ADDRESS", key: "email", width: 25 }, // id="email"
  { header: "SPOUSE SURNAME", key: "spouseSurname", width: 20 }, // id="spouseSurname"
  { header: "SPOUSE FIRST NAME", key: "spouseFirstname", width: 20 }, // id="spouseFirstname"
  { header: "SPOUSE MIDDLE NAME", key: "spouseMiddlename", width: 20 }, // id="spouseMiddlename"
  { header: "SPOUSE EXTENSION", key: "spouseExtension", width: 15 }, // id="spouseExtension"
  { header: "SPOUSE OCCUPATION", key: "spouseOccupation", width: 20 }, // id="spouseOccupation"
  { header: "CHILD NAME 1", key: "childName1", width: 20 }, // id="childName1"
  { header: "CHILD DOB 1", key: "childDob1", width: 15 }, // id="childDob1"
  { header: "FATHER SURNAME", key: "fatherSurname", width: 20 }, // id="fatherSurname"
  { header: "FATHER FIRST NAME", key: "fatherFirstname", width: 20 }, // id="fatherFirstname"
  { header: "FATHER MIDDLE NAME", key: "fatherMiddlename", width: 20 }, // id="fatherMiddlename"
  { header: "FATHER EXTENSION", key: "fatherExtension", width: 15 }, // id="fatherExtension"
  { header: "FATHER OTHER", key: "fatherOther", width: 15 }, // id="fatherOther"
  { header: "FATHER DOB", key: "fatherDob", width: 15 }, // id="fatherDob"
  { header: "MOTHER SURNAME", key: "motherSurname", width: 20 }, // id="motherSurname"
  { header: "MOTHER FIRST NAME", key: "motherFirstname", width: 20 }, // id="motherFirstname"
  { header: "MOTHER MIDDLE NAME", key: "motherMiddlename", width: 20 }, // id="motherMiddlename"
  { header: "MOTHER OTHER", key: "motherOther", width: 15 }, // id="motherOther"
  { header: "MOTHER DOB", key: "motherDob", width: 15 }, // id="motherDob"
  { header: "ELEMENTARY SCHOOL", key: "elemSchool", width: 25 }, // id="elemSchool"
  { header: "ELEMENTARY DEGREE", key: "elemDegree", width: 20 }, // id="elemDegree"
  { header: "ELEMENTARY FROM", key: "elemFrom", width: 10 }, // id="elemFrom"
  { header: "ELEMENTARY TO", key: "elemTo", width: 10 }, // id="elemTo"
  { header: "ELEMENTARY UNITS", key: "elemUnits", width: 10 }, // id="elemUnits"
  { header: "ELEMENTARY YEAR GRADUATED", key: "elemYearGrad", width: 15 }, // id="elemYearGrad"
  { header: "ELEMENTARY HONORS", key: "elemHonors", width: 20 }, // id="elemHonors"
  { header: "SECONDARY SCHOOL", key: "secSchool", width: 25 }, // id="secSchool"
  { header: "SECONDARY DEGREE", key: "secDegree", width: 20 }, // id="secDegree"
  { header: "SECONDARY FROM", key: "secFrom", width: 10 }, // id="secFrom"
  { header: "SECONDARY TO", key: "secTo", width: 10 }, // id="secTo"
  { header: "SECONDARY UNITS", key: "secUnits", width: 10 }, // id="secUnits"
  { header: "SECONDARY YEAR GRADUATED", key: "secYearGrad", width: 15 }, // id="secYearGrad"
  { header: "SECONDARY HONORS", key: "secHonors", width: 20 }, // id="secHonors"
  { header: "VOCATIONAL SCHOOL", key: "vocSchool", width: 25 }, // id="vocSchool"
  { header: "VOCATIONAL DEGREE", key: "vocDegree", width: 20 }, // id="vocDegree"
  { header: "VOCATIONAL FROM", key: "vocFrom", width: 10 }, // id="vocFrom"
  { header: "VOCATIONAL TO", key: "vocTo", width: 10 }, // id="vocTo"
  { header: "VOCATIONAL UNITS", key: "vocUnits", width: 10 }, // id="vocUnits"
  { header: "VOCATIONAL YEAR GRADUATED", key: "vocYearGrad", width: 15 }, // id="vocYearGrad"
  { header: "VOCATIONAL HONORS", key: "vocHonors", width: 20 }, // id="vocHonors"
  { header: "COLLEGE SCHOOL", key: "colSchool", width: 25 }, // id="colSchool"
  { header: "COLLEGE DEGREE", key: "colDegree", width: 20 }, // id="colDegree"
  { header: "COLLEGE FROM", key: "colFrom", width: 10 }, // id="colFrom"
  { header: "COLLEGE TO", key: "colTo", width: 10 }, // id="colTo"
  { header: "COLLEGE UNITS", key: "colUnits", width: 10 }, // id="colUnits"
  { header: "COLLEGE YEAR GRADUATED", key: "colYearGrad", width: 15 }, // id="colYearGrad"
  { header: "COLLEGE HONORS", key: "colHonors", width: 20 }, // id="colHonors"
  { header: "GRADUATE SCHOOL", key: "gradSchool", width: 25 }, // id="gradSchool"
  { header: "GRADUATE DEGREE", key: "gradDegree", width: 20 }, // id="gradDegree"
  { header: "GRADUATE FROM", key: "gradFrom", width: 10 }, // id="gradFrom"
  { header: "GRADUATE TO", key: "gradTo", width: 10 }, // id="gradTo"
  { header: "GRADUATE UNITS", key: "gradUnits", width: 10 }, // id="gradUnits"
  { header: "GRADUATE YEAR GRADUATED", key: "gradYearGrad", width: 15 }, // id="gradYearGrad"
  { header: "GRADUATE HONORS", key: "gradHonors", width: 20 }, // id="gradHonors"

  //page 2
  { header: "Service 1", key: "service1", width: 55 },
  { header: "Rating 1", key: "rating1", width: 20 },
  { header: "Exam Date 1", key: "examDate1", width: 25 },
  { header: "Exam Place 1", key: "examPlace1", width: 50 },
  { header: "License No 1", key: "licenseNo1", width: 65 },
  { header: "License Validity 1", key: "licenseValidity1", width: 35 },
  { header: "Service 2", key: "service2", width: 55 },
  { header: "Rating 2", key: "rating2", width: 20 },
  { header: "Exam Date 2", key: "examDate2", width: 25 },
  { header: "Exam Place 2", key: "examPlace2", width: 50 },
  { header: "License No 2", key: "licenseNo2", width: 65 },
  { header: "License Validity 2", key: "licenseValidity2", width: 15 },
  { header: "Service 3", key: "service3", width: 55 },
  { header: "Rating 3", key: "rating3", width: 20 },
  { header: "Exam Date 3", key: "examDate3", width: 25 },
  { header: "Exam Place 3", key: "examPlace3", width: 50 },
  { header: "License No 3", key: "licenseNo3", width: 65 },
  { header: "License Validity 3", key: "licenseValidity3", width: 35 },
  { header: "Service 4", key: "service4", width: 55 },
  { header: "Rating 4", key: "rating4", width: 20 },
  { header: "Exam Date 4", key: "examDate4", width: 25 },
  { header: "Exam Place 4", key: "examPlace4", width: 50 },
  { header: "License No 4", key: "licenseNo4", width: 65 },
  { header: "License Validity 4", key: "licenseValidity4", width: 35 },
  { header: "Service 5", key: "service5", width: 55 },
  { header: "Rating 5", key: "rating5", width: 20 },
  { header: "Exam Date 5", key: "examDate5", width: 15 },
  { header: "Exam Place 5", key: "examPlace5", width: 20 },
  { header: "License No 5", key: "licenseNo5", width: 65 },
  { header: "License Validity 5", key: "licenseValidity5", width: 35 },
  { header: "Service 6", key: "service6", width: 55 },
  { header: "Rating 6", key: "rating6", width: 20 },
  { header: "Exam Date 6", key: "examDate6", width: 25 },
  { header: "Exam Place 6", key: "examPlace6", width: 50 },
  { header: "License No 6", key: "licenseNo6", width: 65 },
  { header: "License Validity 6", key: "licenseValidity6", width: 35 },
  // WORK EXPERIENCE (10 rows)
  { header: "From 1", key: "from1", width: 15 },
  { header: "To 1", key: "to1", width: 15 },
  { header: "Position 1", key: "position1", width: 20 },
  { header: "Department 1", key: "department1", width: 30 },
  { header: "Salary 1", key: "salary1", width: 15 },
  { header: "GradeStep 1", key: "gradeStep1", width: 15 },
  { header: "AppointmentStatus 1", key: "appointmentStatus1", width: 20 },
  { header: "GovService 1", key: "govService1", width: 10 },
  { header: "From 2", key: "from2", width: 15 },
  { header: "To 2", key: "to2", width: 15 },
  { header: "Position 2", key: "position2", width: 20 },
  { header: "Department 2", key: "department2", width: 30 },
  { header: "Salary 2", key: "salary2", width: 15 },
  { header: "GradeStep 2", key: "gradeStep2", width: 15 },
  { header: "AppointmentStatus 2", key: "appointmentStatus2", width: 20 },
  { header: "GovService 2", key: "govService2", width: 10 },
  { header: "From 3", key: "from3", width: 15 },
  { header: "To 3", key: "to3", width: 15 },
  { header: "Position 3", key: "position3", width: 20 },
  { header: "Department 3", key: "department3", width: 30 },
  { header: "Salary 3", key: "salary3", width: 15 },
  { header: "GradeStep 3", key: "gradeStep3", width: 15 },
  { header: "AppointmentStatus 3", key: "appointmentStatus3", width: 20 },
  { header: "GovService 3", key: "govService3", width: 10 },
  { header: "From 4", key: "from4", width: 15 },
  { header: "To 4", key: "to4", width: 15 },
  { header: "Position 4", key: "position4", width: 20 },
  { header: "Department 4", key: "department4", width: 30 },
  { header: "Salary 4", key: "salary4", width: 15 },
  { header: "GradeStep 4", key: "gradeStep4", width: 15 },
  { header: "AppointmentStatus 4", key: "appointmentStatus4", width: 20 },
  { header: "GovService 4", key: "govService4", width: 10 },
  { header: "From 5", key: "from5", width: 15 },
  { header: "To 5", key: "to5", width: 15 },
  { header: "Position 5", key: "position5", width: 20 },
  { header: "Department 5", key: "department5", width: 30 },
  { header: "Salary 5", key: "salary5", width: 15 },
  { header: "GradeStep 5", key: "gradeStep5", width: 15 },
  { header: "AppointmentStatus 5", key: "appointmentStatus5", width: 20 },
  { header: "GovService 5", key: "govService5", width: 10 },
  { header: "From 6", key: "from6", width: 15 },
  { header: "To 6", key: "to6", width: 15 },
  { header: "Position 6", key: "position6", width: 20 },
  { header: "Department 6", key: "department6", width: 30 },
  { header: "Salary 6", key: "salary6", width: 15 },
  { header: "GradeStep 6", key: "gradeStep6", width: 15 },
  { header: "AppointmentStatus 6", key: "appointmentStatus6", width: 20 },
  { header: "GovService 6", key: "govService6", width: 10 },
  { header: "From 7", key: "from7", width: 15 },
  { header: "To 7", key: "to7", width: 15 },
  { header: "Position 7", key: "position7", width: 20 },
  { header: "Department 7", key: "department7", width: 30 },
  { header: "Salary 7", key: "salary7", width: 15 },
  { header: "GradeStep 7", key: "gradeStep7", width: 15 },
  { header: "AppointmentStatus 7", key: "appointmentStatus7", width: 20 },
  { header: "GovService 7", key: "govService7", width: 10 },
  { header: "From 8", key: "from8", width: 15 },
  { header: "To 8", key: "to8", width: 15 },
  { header: "Position 8", key: "position8", width: 20 },
  { header: "Department 8", key: "department8", width: 30 },
  { header: "Salary 8", key: "salary8", width: 15 },
  { header: "GradeStep 8", key: "gradeStep8", width: 15 },
  { header: "AppointmentStatus 8", key: "appointmentStatus8", width: 20 },
  { header: "GovService 8", key: "govService8", width: 10 },
  { header: "From 9", key: "from9", width: 15 },
  { header: "To 9", key: "to9", width: 15 },
  { header: "Position 9", key: "position9", width: 20 },
  { header: "Department 9", key: "department9", width: 30 },
  { header: "Salary 9", key: "salary9", width: 15 },
  { header: "GradeStep 9", key: "gradeStep9", width: 15 },
  { header: "AppointmentStatus 9", key: "appointmentStatus9", width: 20 },
  { header: "GovService 9", key: "govService9", width: 10 },
  { header: "From 10", key: "from10", width: 15 },
  { header: "To 10", key: "to10", width: 15 },
  { header: "Position 10", key: "position10", width: 20 },
  { header: "Department 10", key: "department10", width: 30 },
  { header: "Salary 10", key: "salary10", width: 15 },
  { header: "GradeStep 10", key: "gradeStep10", width: 15 },
  { header: "AppointmentStatus 10", key: "appointmentStatus10", width: 20 },
  { header: "GovService 10", key: "govService10", width: 10 },

  //page 3 can be added similarly
  // VOLUNTARY WORK (10 rows)
  { header: "Vol Org 1", key: "volOrg1", width: 40 },
  { header: "Vol From 1", key: "volFrom1", width: 15 },
  { header: "Vol To 1", key: "volTo1", width: 15 },
  { header: "Vol Hours 1", key: "volHours1", width: 12 },
  { header: "Vol Position 1", key: "volPosition1", width: 26 },
  { header: "Vol Org 2", key: "volOrg2", width: 40 },
  { header: "Vol From 2", key: "volFrom2", width: 15 },
  { header: "Vol To 2", key: "volTo2", width: 15 },
  { header: "Vol Hours 2", key: "volHours2", width: 12 },
  { header: "Vol Position 2", key: "volPosition2", width: 26 },
  { header: "Vol Org 3", key: "volOrg3", width: 40 },
  { header: "Vol From 3", key: "volFrom3", width: 15 },
  { header: "Vol To 3", key: "volTo3", width: 15 },
  { header: "Vol Hours 3", key: "volHours3", width: 12 },
  { header: "Vol Position 3", key: "volPosition3", width: 26 },
  { header: "Vol Org 4", key: "volOrg4", width: 40 },
  { header: "Vol From 4", key: "volFrom4", width: 15 },
  { header: "Vol To 4", key: "volTo4", width: 15 },
  { header: "Vol Hours 4", key: "volHours4", width: 12 },
  { header: "Vol Position 4", key: "volPosition4", width: 26 },
  { header: "Vol Org 5", key: "volOrg5", width: 40 },
  { header: "Vol From 5", key: "volFrom5", width: 15 },
  { header: "Vol To 5", key: "volTo5", width: 15 },
  { header: "Vol Hours 5", key: "volHours5", width: 12 },
  { header: "Vol Position 5", key: "volPosition5", width: 26 },
  { header: "Vol Org 6", key: "volOrg6", width: 40 },
  { header: "Vol From 6", key: "volFrom6", width: 15 },
  { header: "Vol To 6", key: "volTo6", width: 15 },
  { header: "Vol Hours 6", key: "volHours6", width: 12 },
  { header: "Vol Position 6", key: "volPosition6", width: 26 },
  { header: "Vol Org 7", key: "volOrg7", width: 40 },
  { header: "Vol From 7", key: "volFrom7", width: 15 },
  { header: "Vol To 7", key: "volTo7", width: 15 },
  { header: "Vol Hours 7", key: "volHours7", width: 12 },
  { header: "Vol Position 7", key: "volPosition7", width: 26 },
  { header: "Vol Org 8", key: "volOrg8", width: 40 },
  { header: "Vol From 8", key: "volFrom8", width: 15 },
  { header: "Vol To 8", key: "volTo8", width: 15 },
  { header: "Vol Hours 8", key: "volHours8", width: 12 },
  { header: "Vol Position 8", key: "volPosition8", width: 26 },
  { header: "Vol Org 9", key: "volOrg9", width: 40 },
  { header: "Vol From 9", key: "volFrom9", width: 15 },
  { header: "Vol To 9", key: "volTo9", width: 15 },
  { header: "Vol Hours 9", key: "volHours9", width: 12 },
  { header: "Vol Position 9", key: "volPosition9", width: 26 },
  { header: "Vol Org 10", key: "volOrg10", width: 40 },
  { header: "Vol From 10", key: "volFrom10", width: 15 },
  { header: "Vol To 10", key: "volTo10", width: 15 },
  { header: "Vol Hours 10", key: "volHours10", width: 12 },
  { header: "Vol Position 10", key: "volPosition10", width: 26 },

  // Learning & Development (10 rows)
  { header: "LD Title 1", key: "ldTitle1", width: 36 },
  { header: "LD From 1", key: "ldFrom1", width: 15 },
  { header: "LD To 1", key: "ldTo1", width: 15 },
  { header: "LD Hours 1", key: "ldHours1", width: 12 },
  { header: "LD Type 1", key: "ldType1", width: 15 },
  { header: "LD Sponsor 1", key: "ldSponsor1", width: 15 },
  { header: "LD Title 2", key: "ldTitle2", width: 36 },
  { header: "LD From 2", key: "ldFrom2", width: 15 },
  { header: "LD To 2", key: "ldTo2", width: 15 },
  { header: "LD Hours 2", key: "ldHours2", width: 12 },
  { header: "LD Type 2", key: "ldType2", width: 15 },
  { header: "LD Sponsor 2", key: "ldSponsor2", width: 15 },
  { header: "LD Title 3", key: "ldTitle3", width: 36 },
  { header: "LD From 3", key: "ldFrom3", width: 15 },
  { header: "LD To 3", key: "ldTo3", width: 15 },
  { header: "LD Hours 3", key: "ldHours3", width: 12 },
  { header: "LD Type 3", key: "ldType3", width: 15 },
  { header: "LD Sponsor 3", key: "ldSponsor3", width: 15 },
  { header: "LD Title 4", key: "ldTitle4", width: 36 },
  { header: "LD From 4", key: "ldFrom4", width: 15 },
  { header: "LD To 4", key: "ldTo4", width: 15 },
  { header: "LD Hours 4", key: "ldHours4", width: 12 },
  { header: "LD Type 4", key: "ldType4", width: 15 },
  { header: "LD Sponsor 4", key: "ldSponsor4", width: 15 },
  { header: "LD Title 5", key: "ldTitle5", width: 36 },
  { header: "LD From 5", key: "ldFrom5", width: 15 },
  { header: "LD To 5", key: "ldTo5", width: 15 },
  { header: "LD Hours 5", key: "ldHours5", width: 12 },
  { header: "LD Type 5", key: "ldType5", width: 15 },
  { header: "LD Sponsor 5", key: "ldSponsor5", width: 15 },
  { header: "LD Title 6", key: "ldTitle6", width: 36 },
  { header: "LD From 6", key: "ldFrom6", width: 15 },
  { header: "LD To 6", key: "ldTo6", width: 15 },
  { header: "LD Hours 6", key: "ldHours6", width: 12 },
  { header: "LD Type 6", key: "ldType6", width: 15 },
  { header: "LD Sponsor 6", key: "ldSponsor6", width: 15 },
  { header: "LD Title 7", key: "ldTitle7", width: 36 },
  { header: "LD From 7", key: "ldFrom7", width: 15 },
  { header: "LD To 7", key: "ldTo7", width: 15 },
  { header: "LD Hours 7", key: "ldHours7", width: 12 },
  { header: "LD Type 7", key: "ldType7", width: 15 },
  { header: "LD Sponsor 7", key: "ldSponsor7", width: 15 },
  { header: "LD Title 8", key: "ldTitle8", width: 36 },
  { header: "LD From 8", key: "ldFrom8", width: 15 },
  { header: "LD To 8", key: "ldTo8", width: 15 },
  { header: "LD Hours 8", key: "ldHours8", width: 12 },
  { header: "LD Type 8", key: "ldType8", width: 15 },
  { header: "LD Sponsor 8", key: "ldSponsor8", width: 15 },
  { header: "LD Title 9", key: "ldTitle9", width: 36 },
  { header: "LD From 9", key: "ldFrom9", width: 15 },
  { header: "LD To 9", key: "ldTo9", width: 15 },
  { header: "LD Hours 9", key: "ldHours9", width: 12 },
  { header: "LD Type 9", key: "ldType9", width: 15 },
  { header: "LD Sponsor 9", key: "ldSponsor9", width: 15 },
  { header: "LD Title 10", key: "ldTitle10", width: 36 },
  { header: "LD From 10", key: "ldFrom10", width: 15 },
  { header: "LD To 10", key: "ldTo10", width: 15 },
  { header: "LD Hours 10", key: "ldHours10", width: 12 },
  { header: "LD Type 10", key: "ldType10", width: 15 },
  { header: "LD Sponsor 10", key: "ldSponsor10", width: 15 },

  // Other Information (10 rows)
  { header: "Skills 1", key: "skills1", width: 33 },
  { header: "Distinction 1", key: "distinction1", width: 33 },
  { header: "Membership 1", key: "membership1", width: 34 },
  { header: "Skills 2", key: "skills2", width: 33 },
  { header: "Distinction 2", key: "distinction2", width: 33 },
  { header: "Membership 2", key: "membership2", width: 34 },
  { header: "Skills 3", key: "skills3", width: 33 },
  { header: "Distinction 3", key: "distinction3", width: 33 },
  { header: "Membership 3", key: "membership3", width: 34 },
  { header: "Skills 4", key: "skills4", width: 33 },
  { header: "Distinction 4", key: "distinction4", width: 33 },
  { header: "Membership 4", key: "membership4", width: 34 },
  { header: "Skills 5", key: "skills5", width: 33 },
  { header: "Distinction 5", key: "distinction5", width: 33 },
  { header: "Membership 5", key: "membership5", width: 34 },
  { header: "Skills 6", key: "skills6", width: 33 },
  { header: "Distinction 6", key: "distinction6", width: 33 },
  { header: "Membership 6", key: "membership6", width: 34 },
  { header: "Skills 7", key: "skills7", width: 33 },
  { header: "Distinction 7", key: "distinction7", width: 33 },
  { header: "Membership 7", key: "membership7", width: 34 },
  { header: "Skills 8", key: "skills8", width: 33 },
  { header: "Distinction 8", key: "distinction8", width: 33 },
  { header: "Membership 8", key: "membership8", width: 34 },
  { header: "Skills 9", key: "skills9", width: 33 },
  { header: "Distinction 9", key: "distinction9", width: 33 },
  { header: "Membership 9", key: "membership9", width: 34 },
  { header: "Skills 10", key: "skills10", width: 33 },
  { header: "Distinction 10", key: "distinction10", width: 33 },
  { header: "Membership 10", key: "membership10", width: 34 },

    // PAGE 4 QUESTIONS
  { header: "Q34a: Related within third degree?", key: "q34a", width: 18 },
  { header: "Q34b: Related within fourth degree?", key: "q34b", width: 18 },
  { header: "Q34 Details", key: "q34_details", width: 30 },
  { header: "Q35a: Found guilty of admin offense?", key: "q35a", width: 18 },
  { header: "Q35a Details", key: "q35a_details", width: 30 },
  { header: "Q35b: Criminally charged?", key: "q35b", width: 18 },
  { header: "Q35b Details", key: "q35b_details", width: 30 },
  { header: "Q35b Date Filed", key: "q35b_date", width: 15 },
  { header: "Q35b Status", key: "q35b_status", width: 20 },
  { header: "Q36: Convicted of crime?", key: "q36", width: 18 },
  { header: "Q36 Details", key: "q36_details", width: 30 },
  { header: "Q37: Separated from service?", key: "q37", width: 18 },
  { header: "Q37 Details", key: "q37_details", width: 30 },
  { header: "Q38a: Candidate in election?", key: "q38a", width: 18 },
  { header: "Q38a Details", key: "q38a_details", width: 30 },
  { header: "Q38b: Resigned to campaign?", key: "q38b", width: 18 },
  { header: "Q38b Details", key: "q38b_details", width: 30 },
  { header: "Q39: Resigned before election?", key: "q39", width: 18 },
  { header: "Q39 Details", key: "q39_details", width: 30 },
  { header: "Q39 Immigrant/Resident?", key: "q39imm", width: 18 },
  { header: "Q39 Immigrant Details", key: "q39imm_details", width: 30 },
  { header: "Q40a: Indigenous group member?", key: "q40a", width: 18 },
  { header: "Q40a Details", key: "q40a_details", width: 30 },
  { header: "Q40b: Person with disability?", key: "q40b", width: 18 },
  { header: "Q40b Details", key: "q40b_details", width: 30 },
  { header: "Q40c: Solo parent?", key: "q40c", width: 18 },
  { header: "Q40c Details", key: "q40c_details", width: 30 },
  
  // REFERENCES
  { header: "Reference 1 Name", key: "refName1", width: 25 },
  { header: "Reference 1 Address", key: "refAddress1", width: 35 },
  { header: "Reference 1 Tel", key: "refTel1", width: 18 },
  { header: "Reference 2 Name", key: "refName2", width: 25 },
  { header: "Reference 2 Address", key: "refAddress2", width: 35 },
  { header: "Reference 2 Tel", key: "refTel2", width: 18 },
  { header: "Reference 3 Name", key: "refName3", width: 25 },
  { header: "Reference 3 Address", key: "refAddress3", width: 35 },
  { header: "Reference 3 Tel", key: "refTel3", width: 18 },
  
  // GOVERNMENT ID
  { header: "Govt ID Type", key: "govtIdType", width: 20 },
  { header: "Govt ID Number", key: "govtIdNumber", width: 20 },
  { header: "Govt ID Issue", key: "govtIdIssue", width: 30 },
];


async function initWorkbook() {
  if (!fs.existsSync(filePath)) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("PersonalDataSheet");

    sheet.columns = columns;

    await workbook.xlsx.writeFile(filePath);
    console.log("Created new Excel file:", filePath);
  }
}



// Save data to Excel
app.post("/save", async (req, res) => {
  try {
    const data = req.body;
    const workbook = new ExcelJS.Workbook();
    let sheet;

    if (fs.existsSync(filePath)) {
      await workbook.xlsx.readFile(filePath);
      sheet = workbook.getWorksheet("PersonalDataSheet");
      if (!sheet) {
        sheet = workbook.addWorksheet("PersonalDataSheet");
        sheet.columns = columns;
      }
    } else {
      sheet = workbook.addWorksheet("PersonalDataSheet");
      sheet.columns = columns;
    }

    // Always set columns before adding rows
    sheet.columns = columns;

    // Find the first empty row
    const row = getFirstEmptyRow(sheet);

    // Fill the row with data
    columns.forEach((col, idx) => {
      row.getCell(idx + 1).value = data[col.key] || "";
    });

    row.commit();
    await workbook.xlsx.writeFile(filePath);

    // Just return success if no error
    res.json({
      message: "Data saved successfully!",
      success: true,
    });
  } catch (err) {
    console.error("Save error:", err);
    res.status(500).json({
      error: "Failed to save data: " + err.message,
      success: false,
    });
  }
});
// Save Page 2 data and merge with Page 1
app.post("/savePage2", async (req, res) => {
  try {
    const { eligibility, workExperience } = req.body;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    let sheet = workbook.getWorksheet("PersonalDataSheet");
    if (!sheet) {
      sheet = workbook.addWorksheet("PersonalDataSheet");
      sheet.columns = columns;
    }

    // Prepare Page 2 data as a flat object
    const page2Data = {};
    eligibility.forEach((row, i) => {
      page2Data[`service${i + 1}`] = row.service || "";
      page2Data[`rating${i + 1}`] = row.rating || "";
      page2Data[`examDate${i + 1}`] = row.examDate || "";
      page2Data[`examPlace${i + 1}`] = row.examPlace || "";
      page2Data[`licenseNo${i + 1}`] = row.licenseNo || "";
      page2Data[`licenseValidity${i + 1}`] = row.licenseValidity || "";
    });
    workExperience.forEach((row, i) => {
      page2Data[`from${i + 1}`] = row.from || "";
      page2Data[`to${i + 1}`] = row.to || "";
      page2Data[`position${i + 1}`] = row.position || "";
      page2Data[`department${i + 1}`] = row.department || "";
      page2Data[`salary${i + 1}`] = row.salary || "";
      page2Data[`gradeStep${i + 1}`] = row.gradeStep || "";
      page2Data[`appointmentStatus${i + 1}`] = row.appointmentStatus || "";
      page2Data[`govService${i + 1}`] = row.govService || "";
    });

    // Find the last row (assume it's the one to update)
    const lastRow = sheet.lastRow;
    if (!lastRow)
      throw new Error("No existing row to update. Please save Page 1 first.");

    // Update last row with Page 2 data using column index
    columns.forEach((col, idx) => {
      if (page2Data.hasOwnProperty(col.key)) {
        lastRow.getCell(idx + 1).value = page2Data[col.key]; // idx+1 because columns are 1-based
      }
    });

    await workbook.xlsx.writeFile(filePath);

    res.json({
      message: "Data saved successfully!",
      success: true,
    });
  } catch (err) {
    console.error("Save error:", err);
    res
      .status(500)
      .json({ error: "Failed to save data: " + err.message, success: false });
  }
});

app.post("/savePage3", async (req, res) => {
  try {
    const { voluntary, ld, otherInfo } = req.body;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    let sheet = workbook.getWorksheet("PersonalDataSheet");
    if (!sheet) {
      sheet = workbook.addWorksheet("PersonalDataSheet");
      sheet.columns = columns;
    }

    // Prepare Page 3 data as a flat object
    const page3Data = {};
    // Voluntary Work
    voluntary.forEach((row, i) => {
      page3Data[`volOrg${i + 1}`] = row.org || "";
      page3Data[`volFrom${i + 1}`] = row.from || "";
      page3Data[`volTo${i + 1}`] = row.to || "";
      page3Data[`volHours${i + 1}`] = row.hours || "";
      page3Data[`volPosition${i + 1}`] = row.position || "";
    });
    // Learning & Development
    ld.forEach((row, i) => {
      page3Data[`ldTitle${i + 1}`] = row.title || "";
      page3Data[`ldFrom${i + 1}`] = row.from || "";
      page3Data[`ldTo${i + 1}`] = row.to || "";
      page3Data[`ldHours${i + 1}`] = row.hours || "";
      page3Data[`ldType${i + 1}`] = row.type || "";
      page3Data[`ldSponsor${i + 1}`] = row.sponsor || "";
    });
    // Other Information
    otherInfo.forEach((row, i) => {
      page3Data[`skills${i + 1}`] = row.skills || "";
      page3Data[`distinction${i + 1}`] = row.distinction || "";
      page3Data[`membership${i + 1}`] = row.membership || "";
    });

    // Find the last row (assume it's the one to update)
    const lastRow = sheet.lastRow;
    if (!lastRow)
      throw new Error("No existing row to update. Please save Page 1 first.");

    // Update last row with Page 3 data using column index
    columns.forEach((col, idx) => {
      if (page3Data.hasOwnProperty(col.key)) {
        lastRow.getCell(idx + 1).value = page3Data[col.key];
      }
    });

    await workbook.xlsx.writeFile(filePath);

    res.json({
      message: "Data saved successfully!",
      success: true,
    });
  } catch (err) {
    console.error("Save error:", err);
    res
      .status(500)
      .json({ error: "Failed to save data: " + err.message, success: false });
  }
});

app.post("/savePage4", async (req, res) => {
  try {
    const data = req.body;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    let sheet = workbook.getWorksheet("PersonalDataSheet");
    if (!sheet) {
      sheet = workbook.addWorksheet("PersonalDataSheet");
      sheet.columns = columns;
    }

    // Find the last row (assume it's the one to update)
    const lastRow = sheet.lastRow;
    if (!lastRow)
      throw new Error("No existing row to update. Please save Page 1 first.");

    // Update last row with Page 4 data using column index
    columns.forEach((col, idx) => {
      if (data.hasOwnProperty(col.key)) {
        lastRow.getCell(idx + 1).value = data[col.key];
      }
    });

    await workbook.xlsx.writeFile(filePath);

    res.json({
      message: "Data saved successfully!",
      success: true,
    });
  } catch (err) {
    console.error("Save error:", err);
    res
      .status(500)
      .json({ error: "Failed to save data: " + err.message, success: false });
  }
});

function getFirstEmptyRow(sheet) {
  for (let i = 2; i <= sheet.rowCount + 1; i++) { // Start at 2 (after header)
    const row = sheet.getRow(i);
    // Check if all cells in the row are empty
    const isEmpty = columns.every((col, idx) => {
      const cell = row.getCell(idx + 1);
      return !cell.value || cell.value === "";
    });
    if (isEmpty) return row;
  }
  // If no empty row, return a new row at the end
  return sheet.getRow(sheet.rowCount + 1);
}

// Start server
app.listen(PORT, async () => {
  await initWorkbook();
  console.log(`✅ Server running at http://localhost:${PORT}`);
});
