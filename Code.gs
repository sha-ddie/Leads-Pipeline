// Reading data sheets 
const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Initial Data");
const ss_1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contacted Leads");
const ss_2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interactions Log");
const ss_3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Website Data");
// const last_row = ss.getLastRow()
const headers = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
const headers_1= ss_1.getRange(1, 1, 1, ss_1.getLastColumn()).getValues()[0];
const headers_2= ss_2.getRange(1, 1, 1, ss_2.getLastColumn()).getValues()[0];
// const index_phone = headers.indexOf("phone_number");


function doGet() {
  const  cust_info  = dataInfo();

  const template = HtmlService.createTemplateFromFile("index");
  template.cust_data = cust_info.length > 0 ? cust_info[0] : {};
  return template.evaluate()
                     .setTitle("Lead Manager");
}

function LASTROW(sheet, columnLetter = "A") {
  const column = sheet.getRange(columnLetter + ":" + columnLetter).getValues();
  // Logger.log("Column length: "+column.length);
  for (let i = column.length-1 ; i >= 0; i--) {
    if ( (column[i][0] !== "") && (column[i+2][0] !== "")) {
      // Logger.log("Last row: "+(i ))
      return i + 3; // Convert 0-based index to 1-based row number
    }
  }
  return 0; // No data found
}
// function getLastRow(){
//   var last = LASTROW(ss);
//   last = last!==0? last:1 ;
//   Logger.log(last)
// }

// Initial tab Function search
function dataInfo(target_no, tab){ 
  // target_no = "720310379"; tab = "initial";
  if(!target_no ) return{};
  if(!tab ) return{};
  console.log("Phone Search: "+tab+" ("+target_no+")");

  if (tab=="initial"){
    const index_phone = headers.indexOf("phone_number");      
    const lastRow = LASTROW(ss);
    if(lastRow < 2) return {}; // no data

    // Logger.log(index_phone+" : (row)"+lastRow)
    var phone_nos = ss.getRange(1,index_phone+1,lastRow,1).getValues();
    // Convert 2D array → 1D array
    var list = phone_nos.map(r => r[0].toString().slice(-9));

    var rowIndex = list.findIndex(x => x === target_no) + 1;
    if(rowIndex === -1||rowIndex === 0) {
      throw new Error("Phone no. not found!!! try another") // not found
      };

    var row = ss.getRange(rowIndex,1,1,ss.getLastColumn()).getValues()[0];
    if (row[headers.indexOf("Lead_status")]=== "Contacted"){
      throw new Error("Phone No. already contacted") 
    }

    var info = {"name": row[headers.indexOf("full_name")],
                "email": row[headers.indexOf("email")],
                "platform": row[headers.indexOf("platform")],
                "purpose":row[headers.indexOf("what_is_the_purpose_of_the_loan?")],
                "amount":row[headers.indexOf("how_much_are_you_looking_to_borrow?")],
                "lead_date":row[headers.indexOf("Lead_date")] 
                            ? new Date(row[headers.indexOf("Lead_date")]).toISOString(): ""
    };            
    Logger.log(info)
    return info 
  }

  if ((tab=="qualified") || (tab=="interactions")){ 
    
    const index_phone = headers_1.indexOf("phone_number");
    const lastRow = LASTROW(ss_1);
    if(lastRow < 2) throw new Error ("Phone No not found.."); // no data

    // Logger.log(index_phone+" : (row)"+lastRow)
    var phone_nos = ss_1.getRange(1,index_phone+1,lastRow,1).getValues()
    // Convert 2D array → 1D array
    var list = phone_nos.map(r => r[0].toString().slice(-9));

    var rowIndex = list.findIndex(x => x === target_no) + 1;
    if(rowIndex === -1||rowIndex === 0) {
      throw new Error("Phone No. not contacted...First Initiate then Search") // not found
      };
    // Logger.log("Row: "+rowIndex)

    var row = ss_1.getRange(rowIndex,1,1,ss.getLastColumn()).getValues()[0];

    interact = row[rowIndex,headers_1.indexOf("interaction_status")]
    if(interact ==="End"){throw new Error("Customer interactions was stopped...for assistance contact ADMIN")}

    var info = { "name": safeVal( row[headers_1.indexOf("full_name")]),
                "staff": safeVal(row[headers_1.indexOf("Assigned Staff")]),
                "purpose":safeVal(row[headers_1.indexOf("loan_purpose")]),
                "amount":safeVal(row[headers_1.indexOf("amount")]) ,
                "application":safeVal(row[headers_1.indexOf("application_status")]) ,
                "documents":safeVal(row[headers_1.indexOf("documents_status")]),
                "staff":safeVal(row[headers_1.indexOf("Assigned Staff")]) };
    Logger.log(info)
    return info 
  }

}


function writeData(data, tab) {
  // data is an object like {staff, method, status, followUp, remarks}
  // data={ "phone": "790958085","documents": "Yes", "appication": "Yes","remarks": "remarks"};
  // tab="initial";
  console.log(tab+": "+data);
 
  if (tab == "initial"){

    const lastRow = LASTROW(ss_1)!==0? LASTROW(ss_1):2 ;
    const lastRow_2 = LASTROW(ss_2)!==0? LASTROW(ss_2):1 ;
      if (!data.phone){ throw new Error("Missing phone number in submitted data.");}

    // Get all phone numbers in column where phone is stored
    const phone_index = headers.indexOf("phone_number")+1;
    const phoneValues = ss_1.getRange(2, phone_index, lastRow - 1, 1).getValues().flat();
    const rowIndex = phoneValues.findIndex(p => String(p) === String(data.phone)) + 2;
    if (rowIndex !== 1) throw new Error("Customer number already contacted.");
    Logger.log(" Phone no row: " + rowIndex)

    // Do whatever you want with it, e.g., write to a sheet
    var values = [[ new Date(),data.lead_date , String(data.phone), data.name,  data.email, 
        data.platform, data.loan_type, data.amount,  data.staff,  data.method,  data.status, 
        data.followUp,   data.remarks]];

    data["type"] = "Initial ";
    var values_2 = [[ new Date(), String(data.phone), data.name, data.staff, data.method, data.type,
          data.remarks]]

    ss_1.getRange(lastRow+1,1,1,values[0].length,).setValues(values);
    ss_2.getRange(lastRow_2+1,1,1,values_2[0].length,).setValues(values_2);  

    if (data.followUp=="No"){ss_1.getRange(lastRow-1, headers_1.indexOf("interaction_status")+1).setValue("End") ;}  

    return "Success";
  }
  
  if (tab=="qualified"){
    // const columns = ["documents_status","applcation_status","qualified_remarks"]
    const lastRow = LASTROW(ss_1)!==0? LASTROW(ss_1):2;

    if (lastRow < 2) return "No data in sheet";
    if (!data.phone){ throw new Error("Missing phone number in submitted data.");}
    // return "Success";

    // Get all phone numbers in column where phone is stored
    const phone_index = headers_1.indexOf("phone_number")+1;

    const phoneValues = ss_1.getRange(2, phone_index, lastRow - 1, 1).getValues().flat();
    // Logger.log(phoneValues)
    
    // Find the row with matching phone number
    const rowIndex = phoneValues.findIndex(p => String(p) === String(data.phone)) + 2;
    Logger.log(" Phone no row: " + rowIndex)
    if (rowIndex === 1) throw new Error("Phone number not found.");

    interact = ss_1.getRange(rowIndex,headers_1.indexOf("interaction_status")+1).getValue()
    if (interact ==="End"){throw new Error("Customer interactions was stopped...for assistance contact ADMIN")}

    // // Map data to columns to update
    let updateColumns = ["documents_status", "application_status", "qualified_remarks","interaction_status","Call Status"];
    let updateValues = [data.documents, data.application, data.remarks, data.interact,"Qualified"];

    updateColumns.forEach((colName, i) => {
      const colIndex = headers_1.indexOf(colName) + 1;
      if (colIndex > 0) {
        ss_1.getRange(rowIndex, colIndex).setValue(updateValues[i]);
        }
    });

    return "Success";
  }

  if (tab == "interactions") {
    const lastRow = LASTROW(ss_2)!==0? LASTROW(ss_2):2 ;
    const lastRow_1 = LASTROW(ss_1)!==0? LASTROW(ss_1):2 ;
    Logger.log(lastRow_1+": "+lastRow);
    if (!data.phone){ throw new Error("Missing phone number ...")};
    
    const phone_index = headers_1.indexOf("phone_number")+1;
    const phoneValues = ss_1.getRange(2, phone_index, lastRow_1 - 1, 1).getValues().flat();
        // Find the uncontacted  phone number
    const rowIndex = phoneValues.findIndex(p => String(p) === String(data.phone)) + 2;
    // Logger.log(phoneValues);
    // Logger.log(rowIndex);
    // Logger.log(data.phone);
    if (rowIndex == 1) throw new Error("Customer not previoulsy contacted... Go to INITIAL TAB");

    var values = [[ new Date(), String(data.phone), data.name, data.staff, data.method, data.type,
          data.remarks ]];

    ss_2.getRange(lastRow+1,1,1,values[0].length,).setValues(values);    

    return "Success";

  }
}

function filterLeads(rowLimit,staff_,page, category_,phone){
  // rowLimit=5 ; staff_="Kelvin" ; page="initial" ; category_="Cold"
  console.log("Parameters:", rowLimit,staff_,page, category_,phone);
  if (page=="initial"){
    data = filterInitial(rowLimit)
    return data
  }
  if (page=="qualified"){
    data = filterQualified(staff_, category_)
    return data
  }
  if (page=="interactions"){
    data = filterInteractions(phone)
    return data
  }
}

function filterInitial(index) {
  // console.log("filterLeads called:", index);
  index=5
  if (!index) return [];

  const lastRow = LASTROW(ss);
  const headerCount = ss.getLastColumn();
  const headersRow = headers

  const statusCol = headersRow.indexOf("Lead_status") + 1;
  
  // Read ALL statuses at once
  const statuses = ss.getRange(2, statusCol, lastRow - 1).getValues().flat();
 

  // Collect last N matching rows
  const targetRows = [];
  for (let i = statuses.length - 1; i >= 0 && targetRows.length < index; i--) {
    if (statuses[i] === "Not Contacted") {
      targetRows.push(i + 2);
    }
  }

  if (targetRows.length === 0) return [];
  
  targetRows.reverse();
  Logger.log(targetRows)

  const columns = [
    "full_name","phone_number","email","platform",
    "what_is_the_purpose_of_the_loan?",
    "how_much_are_you_looking_to_borrow?",
    "Lead_date","Lead_status"
  ];

  // const colIndexes = columns.map(c => headersRow.indexOf(c) + 1);
  const phoneIndex = headersRow.indexOf("phone_number") + 1;

  // READ ALL DATA
  const output = targetRows.map((rowNum, i) => {
    const rowData = ss.getRange( rowNum, 1, 1,  headerCount ).getValues()[0];
    // Logger.log(rowData)
    return {
      status: safeVal(rowData[headersRow.indexOf("Lead_status")]),
      date: rowData[headersRow.indexOf("Lead_date")] 
              ? new Date(rowData[headersRow.indexOf("Lead_date")]).toISOString()
              : "",
      phone: rowData[headersRow.indexOf("phone_number")] 
              ? String(rowData[headersRow.indexOf("phone_number")]).slice(-9) 
              : "",
      name: safeVal(rowData[headersRow.indexOf("full_name")]),
      //email: safeVal(rowData[headersRow.indexOf("email")]),
      platform: safeVal(rowData[headersRow.indexOf("platform")]),
      purpose: safeVal(rowData[headersRow.indexOf("what_is_the_purpose_of_the_loan?")]),
      amount: rowData[headersRow.indexOf("how_much_are_you_looking_to_borrow?")] 
              ? String(rowData[headersRow.indexOf("how_much_are_you_looking_to_borrow?")])
              : ""
    };
  });
  // Logger.log(output);
  // Always return an array
  return output || [];
}


function filterQualified(staff, category, limit = 20) {
  // staff="Kelvin";  index=5; category="Cold";
  // console.log("filterQualified:", staff, category, limit);

  if (!staff) throw new Error("Missing staff name.");
  if (!category) throw new Error("Missing category.");

  const lastRow = LASTROW(ss_1);
  if (lastRow<=2)throw new Error("No data found.");
  const headersRow = headers_1;

  const staffCol = headersRow.indexOf("Assigned Staff") + 1;
  const categoryCol = headersRow.indexOf("lead_category") + 1;
  const interactCol = headersRow.indexOf("interaction_status") + 1;
  // Logger.log("staff col: "+ staffCol+" cat col: "+ categoryCol+"interact col: "+interactCol) ;

  // 🚀 Read ONLY the two filter columns
  const staffValues = ss_1.getRange(2, staffCol, lastRow - 1, 1).getValues().flat();
  const catValues   = ss_1.getRange(2, categoryCol, lastRow - 1, 1).getValues().flat();
  const interactValues   = ss_1.getRange(2, interactCol, lastRow - 1, 1).getValues().flat();

  const targetRows = [];
  
  // Scan bottom-up to get most recent first
  for (let i = staffValues.length-1 ; i >= 0 && targetRows.length < limit; i--) {
    // Logger.log(i+": "+staffValues[i]+" : "+catValues[i] +" : "+ interactValues[i])
    if (staffValues[i] === staff && catValues[i] === category && interactValues[i] !== "End") {
      // Logger.log("True")
      targetRows.push(i + 2); // sheet row index
    }
  }
  
  // Logger.log(staffValues);
  if (targetRows.length === 0) return [];

  targetRows.reverse();

  // --------- MAP HEADERS ---------- //
  const col = name => headersRow.indexOf(name) + 1;
  const neededCols = [ col("Assigned Staff"), col("full_name"),col("phone_number"), col("loan_purpose"),
    col("amount"), col("lead_date"), col("Contacted Date"),col("Contact Method"),col("lead_category"),
    col("Call Status")
  ];

  const output = [];

  // 🚀 Read ONLY rows you need, columns you need
  targetRows.forEach(r => {
    const vals = ss_1.getRange(r, 1, 1, headersRow.length).getValues()[0];
    output.push({
      staff: safeVal(vals[col("Assigned Staff") - 1]),
      name: safeVal(vals[col("full_name") - 1]),
      phone: vals[col("phone_number") - 1]
        ? String(vals[col("phone_number") - 1]).slice(-9)
        : "",
      purpose: safeVal(vals[col("loan_purpose") - 1]),
      amount: safeVal(vals[col("amount") - 1]),
      date: vals[col("lead_date") - 1]
        ? new Date(vals[col("lead_date") - 1]).toISOString()
        : "",
      contacted_date: vals[col("Contacted Date") - 1]
        ? new Date(vals[col("Contacted Date") - 1]).toISOString()
        : "",
      //method: safeVal(vals[col("Contact Method") - 1]),
      category: safeVal(vals[col("lead_category") - 1]),
      call_satus: safeVal(vals[col("Call Status") - 1])
    });
  });

  Logger.log(output);
  return output;
}

function filterInteractions( phone, limit = 7) {
  // phone="742227280";
  console.log("filterInteractions:", phone);

  // if (!staff) throw new Error("Missing staff name.");
  if (!phone) throw new Error("Missing phone no.");

  const lastRow = LASTROW(ss_2)!==0? LASTROW(ss_2):2;
  // const lastRow = LASTROW(ss_2);Logger.log(lastRow);
  if (lastRow<1)throw new Error("No data found.");
  const headersRow = headers_2;


  const PhoneCol = headersRow.indexOf("phone_number")+1;

  // 🚀 Read ONLY the two filter columns
  const phoneValues = ss_2.getRange(2, PhoneCol, lastRow , 1).getValues().flat();
  // Logger.log(phoneValues)
  const targetRows = [];

  // Scan bottom-up to get most recent first
  for (let i = phoneValues.length-1 ; i >= 0 && targetRows.length < limit; i--) {
    // Logger.log(i+": "+phoneValues[i]+" : "+phone )
    if (phoneValues[i] === phone ) {
      targetRows.push(i + 2); // sheet row index
    }
  }
  // Logger.log(targetRows)

  if (targetRows.length === 0) return [];

  // targetRows.reverse();

  // --------- MAP HEADERS ---------- //
  const col = name => headersRow.indexOf(name) + 1;
  const neededCols = [col("Date"), col("Assigned Staff"),col("phone_number"), col("Contact Method"),col("Contact Type"),  col("Remarks")];

  const output = [];

  // 🚀 Read ONLY rows you need, columns you need
  targetRows.forEach(r => {
    const vals = ss_2.getRange(r, 1, 1, headersRow.length).getValues()[0];
    output.push({
      date: vals[col("Date") - 1]
        ? new Date(vals[col("Date") - 1]).toISOString()
        : "",
      staff: safeVal(vals[col("Assigned Staff") - 1]),
      phone: vals[col("phone_number") - 1]
        ? String(vals[col("phone_number") - 1]).slice(-9)
        : "",
      method: safeVal(vals[col("Contact Method") - 1]),
      type: safeVal(vals[col("Contact Type") - 1]),
      remarks: safeVal(vals[col("Remarks") - 1])
    });
  });

  Logger.log(output);
  return output;
}


function safeVal(val) {
  return (val === null || val === undefined) ? "" : val;
}

// Hiden section for unqualified to add extar section
function toggleFollowUp(show) {
  document.getElementById("followUpContainer").style.display = show ? "block" : "none";
}


function submitWebsiteData(rawText) {
  // rawText=`
  //   Name
  // Towett Geoffrey
  // Phone Number
  // Email Address
  // tkjeff82@gmail.com
  // Which Loan Are You Interested In?
  // Business Loan
  // Collateral
  // Developed Land
  // Loan Amount
  // 1,500,000
  // Monthly Income Range
  // 20,000 – 50,000
  // Purpose of Loan
  // Expand business
  // Consent
  // checked`
  console.log("Websited data text:  "+rawText);

  // Read header row so we know which fields exist
  const headers_3 = ss_3.getRange(1, 1, 1, ss_3.getLastColumn()).getValues()[0];

  var lastRow = LASTROW(ss_3)!==0 ? LASTROW(ss_3): 2;

    // Get all phone numbers in column where phone is stored
  const phone_index = headers_3.indexOf("Phone Number")+1;

  const phoneValues = ss_3.getRange(2, phone_index, lastRow-1, 1).getValues().flat();

  // Split text into cleaned lines (removes empty lines)
  const lines = rawText
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l.length > 0);
  Logger.log(lines);

  // Convert label/value pairs into an object
  let dataObj = {};
  for (let i = 0; i < lines.length; i += 2) {
    const label = String(lines[i]).trim();
    // Logger.log(label);
    if(!headers_3.includes(label)) throw new Error(`Column name does not exist: "${label}"... Please check preview.`);
    const value = lines[i + 1] || "";
    dataObj[label] = value;
  }
  Logger.log(dataObj);
  dataObj["platform"]="Website";

  var phoneno = String(dataObj["Phone Number"]).replace(/\s+/g, "").slice(-9);
  dataObj["Phone Number"] = phoneno;
  const rowIndex = phoneValues.findIndex(p => String(p) === String(phoneno)) + 1;
  Logger.log(rowIndex)
  if (rowIndex > 0) throw new Error("Data already uploaded..");

  // Construct row in exact same order as headers
  const newRow = headers_3.map((header, index) => {
    if (index === 0) { return new Date();    }// First column = Date timestamp
    return dataObj[header] ||"" ;  // use value if match, else blank
  });
  Logger.log(newRow)

  // Append row
  ss_3.appendRow(newRow);

  return "Data Saves successfuly";
}






