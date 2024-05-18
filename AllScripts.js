const workbook = SpreadsheetApp.getActiveSpreadsheet()
const allSheets = workbook.getSheets();
const wrongData = workbook.getSheetByName('Wrong data');
var ui = SpreadsheetApp.getUi();

function deleteDuplicates () {
  const names = [];
  for (let i = 0; i < allSheets.length; i++ ) {
    const current_sheet = allSheets[i];
    const current_name = allSheets[i].getName();
    const len = current_name.length;
    if (current_name[len-3] === '(' & current_name[len-1] === ')') {
      const nameSliced = current_name.slice(0, len-4);
      if (nameSliced === 'General scripts' || nameSliced === 'Wrong data' || nameSliced === 'Задолженность Отчет' || nameSliced === 'Служебный') {
        ui.alert(`Проверьте лист ${current_name} и лист ${nameSliced} и оставьте один`)
      } else if (names.includes(nameSliced))  {
        const ss = workbook.getSheetByName(nameSliced);
        workbook.deleteSheet(ss);
        current_sheet.setName(nameSliced)
      } else names.push(current_name)
    } else names.push(current_name)
  }
}

function formatDates () {
  for (let sheet of allSheets) {
    if (sheet.getName() !== 'General scripts' & sheet.getName() !=='ТОТАЛ' & sheet.getName() !== 'Wrong data' & sheet.getName() !=='Задолженность Отчет' & sheet.getName() !=='Служебный') {
    let range = sheet.getRange('C8:C');
    range.setNumberFormat('dd.mm.yyyy');
    range = sheet.getRange('V8:V');
    range.setNumberFormat('dd.mm.yyyy');
    range = sheet.getRange('S8:S');
    range.setNumberFormat('dd.mm.yyyy');
    range = sheet.getRange('C3:C6');
    range.setNumberFormat('dd.mm.yyyy');
    range = sheet.getRange('O3:O4');
    range.setNumberFormat('dd.mm.yyyy');
    range = sheet.getRange('R6:R6');
    range.setNumberFormat('dd.mm.yyyy');
    }
  }
}

function checkData () {
  let j = 3;
  let j2 = 3;
  
  const reportRange = wrongData.getRange('A3:E');
  reportRange.clearContent();
  for (let i = 5; i < allSheets.length; i++) {
  
    const sheet = allSheets[i];
    const name = sheet.getName();
    
    const shipDate = sheet.getRange('C7').getCell(1, 1).getValue();
    const amount = sheet.getRange('R7').getCell(1, 1).getValue();
    const dueDate = sheet.getRange('S7').getCell(1, 1).getValue();
    const payment = sheet.getRange('U7').getCell(1, 1).getValue();
    const paymentDate = sheet.getRange('V7').getCell(1, 1).getValue();
    
    if (shipDate !== 'дата') {
      wrongData.getRange(`A${j}`).getCell(1,1).setValue(name)
      wrongData.getRange(`B${j}`).getCell(1,1).setValue(shipDate)
      j++
    }
    if (amount !== 'Итого') {
      wrongData.getRange(`A${j}`).getCell(1,1).setValue(name)
      wrongData.getRange(`B${j}`).getCell(1,1).setValue(amount)
      j++
    }
    if (dueDate !== 'Дата оплаты согласно  договора') {
      wrongData.getRange(`A${j}`).getCell(1,1).setValue(name)
      wrongData.getRange(`B${j}`).getCell(1,1).setValue(dueDate)
      j++
    }
    if (payment !== 'Сумма поступления') {
      wrongData.getRange(`A${j}`).getCell(1,1).setValue(name)
      wrongData.getRange(`B${j}`).getCell(1,1).setValue(payment)
      j++
    }
    if (paymentDate !== 'Дата поступления') {
      wrongData.getRange(`A${j}`).getCell(1,1).setValue(name)
      wrongData.getRange(`B${j}`).getCell(1,1).setValue(paymentDate)
      j++
    }
// check sum 
    const sheetData = sheet.getRange('C8:Q').getValues();
    for (row of sheetData) {
      const date = row[0];
      if (new Date(date) > new Date("1/1/23")) {
        let s = 0;
        for (let i = 0; i < 6; i++) {
          s += Number(row[8 + i]);
        }
        if (s != row[14]) {
          console.log('s', s, 'row14', row[14])
          wrongData.getRange(`D${j2}`).getCell(1,1).setValue(name);
          wrongData.getRange(`E${j2}`).getCell(1,1).setValue(date);
          j2++;
        }
      }
    }

  }
}

function debtReport () {
  const data = {};
  // const dataWorkbook = SpreadsheetApp.openById('1RyCutj4OmTQUKEXa-ULx7eXB7tqgZ0XOl-M8ij7Sx2Q');
  const reportSheet = workbook.getSheetByName('Задолженность отчет');
  const startDate = new Date(reportSheet.getRange('C4').getCell(1,1).getValue());
  const endDate = new Date(reportSheet.getRange('C5').getCell(1,1).getValue());
  const buyer = reportSheet.getRange('B2').getCell(1,1).getValue();
  const contract = reportSheet.getRange('B3').getCell(1,1).getValue();
  let w = 3;
  let w2 = 3;
  let alert = true;
  let errRange = wrongData.getRange('G3:K'); // 
  errRange.clearContent();

  function filterByName (sheet) {
    const name = sheet.getName();
    const idx = name.indexOf('$');
    const companyName = idx === -1 ? name : name.slice(0, idx);
    return buyer === 'ВСЕ' ? true : companyName === buyer
  }

  function filterByContract (sheet) {
    const name = sheet.getName();
    return contract === 'ВСЕ' ? true : name === contract
  }


  for (sheet of workbook.getSheets().filter(filterByName).filter(filterByContract)) {
    // console.log('sheetName',sheet.getName())
    const sheetName = sheet.getName();
    const idx = sheetName.indexOf('$');
    const companyName = idx === -1 ? sheetName : sheetName.slice(0, idx);
    const contractName = idx === -1 ? '' : sheetName.slice(idx+1);
    // console.log('sheet', sheetName, 'idx', idx, 'company', companyName, 'contract', contractName);
    if (sheetName !== 'General scripts' & sheetName !=='ТОТАЛ' & sheetName !== 'Wrong data' & sheetName !=='Задолженность Отчет' & sheetName !=='Служебный') {
      if (!(companyName in data)) {
        data[companyName] = {}
      }
      if (!(contractName in data[companyName])) {
        data[companyName][contractName] = {startDebt : 0,  shipments : 0, payments : 0, endDebt : 0, overdueDebt : 0, debtDays : 0}
      }
      const cur_client = data[companyName][contractName];
    
  
      // start extracting data from sheet 
      const sheet_data = sheet.getRange('C8:V').getValues();
      
      const shipments = [];
      let shipment_start = 0;
      let shipment_end = 0;
      let payment_start = 0;
      let payment_end = 0;
      let have_to_pay = 0;

  

      for (row of sheet_data) {
          // row[0] current date of shipment
          if (row[0]) {
            const curDate = new Date(row[0]);
            // row[16] day of payment by contract
            if (!row[16]) {
              if (alert) {
                alert = false;
                ui.alert(`Внимание! Клиент ${companyName} Договор ${contractName} Дата поставки ${row[0]}\ ДАТА ДОГОВОРА НЕ ЗАПОЛНЕНА Список всех см. на листе Wrong data`);
              }
              errRange = wrongData.getRange(`G${w}:H${w}`);
              errRange.getCell(1,1).setValue(sheetName);
              errRange.getCell(1,2).setValue(row[0]);
              w++;
            }
            const dueDate = new Date(row[16]);
            // row[15] amount of shipment
            if (dueDate <= endDate) {
              have_to_pay += Number(row[15])
              shipments.push({amount : Number(row[15]), duedate : dueDate})
            }

            if (curDate <= endDate) {
              shipment_end += Number(row[15])
            }

            if (curDate < startDate) {
              shipment_start += Number(row[15])
            }
          }
      }

      for (row of sheet_data) {
        // row[19] - payment date
        if (row[19]) {
          const curDate = new Date(row[19]);

          if (curDate <= endDate) {
              payment_end += Number(row[18])
            }

          if (curDate < startDate) {
              payment_start += Number(row[18])
            }
        } else if (row[18]){
          // console.log('row18', row[18], 'start ',payment_start, ' end', payment_end)
          payment_end += Number(row[18]);
          payment_start += Number(row[18]);
          const errR = wrongData.getRange(`J${w2}:K${w2}`);
          errR.getCell(1,1).setValue(sheetName);
          errR.getCell(1,2).setValue(row[18]);
          w2++;
          if (alert) {
            alert = false;
            ui.alert(`Внимание ${sheetName} сумма  ${row[18]}! Дата оплаты  не заполнена, см Wrong data`)
          }
        }
      }
      // add debt on 01.01.18
      let prevDebt = -Number(sheet.getRange('I1').getCell(1,1).getValue());
      // if (sheet.getRange('C8').getCell(1,1).getValue()) {
      //   prevDebt = 0;
      // }
      // console.log(prevDebt)
     
       // console.log('company', companyName,'debt', prevDebt);
      shipments.unshift({duedate : new Date('01.01.2018'), amount : prevDebt})
      have_to_pay += prevDebt;
      shipment_end += prevDebt;
      shipment_start += prevDebt;
      
      const filterDueDate = (date) => (elt) => elt.duedate <= date;

      // console.log('shipments start', shipment_start, 'shipment_end', shipment_end)
      // console.log('payment_start', payment_start, 'payment_end', payment_end)
      cur_client.startDebt = (shipment_start - payment_start);
      cur_client.endDebt = (shipment_end - payment_end);
      cur_client.shipments = (shipment_end - shipment_start);
      cur_client.payments = (payment_end - payment_start);
      cur_client.overdueDebt = (have_to_pay - payment_end)

      // console.log(cashflow)
      // const cashflow_report_end = cashflow.filter(filterDueDate(endDate));

      // console.log(cashflow_report_end.sort((a,b) => b.duedate - a.duedate))

      let targetDate = endDate;
      if (cur_client.overdueDebt > 10 ){
        let sum = cur_client.overdueDebt;
        for (shipment of shipments.reverse()) {
          // console.log('sum', sum, 'amount', shipment.amount, 'due', shipment.duedate)
          sum -= shipment.amount;
          // console.log(sum)
          if (sum <= 10) {
            targetDate = shipment.duedate;
            break;
          }
        } 
      }
      // console.log('debd days?', endDate - targetDate)
      cur_client.debtDays = (endDate - targetDate)/1000/60/60/24;
    } //  if statement end
  } // loop all sheets end

  let j = 1;
  const reportRange = reportSheet.getRange('A9:H');
  reportRange.clearContent();

  for (_buyer in data) {
    for (_contract in data[_buyer]) {
      reportRange.getCell(j, 1).setValue(_buyer);
      reportRange.getCell(j, 2).setValue(_contract);
      reportRange.getCell(j, 3).setValue(data[_buyer][_contract].startDebt);
      reportRange.getCell(j, 4).setValue(data[_buyer][_contract].shipments);
      reportRange.getCell(j, 5).setValue(data[_buyer][_contract].payments);
      reportRange.getCell(j, 6).setValue(data[_buyer][_contract].endDebt);
      reportRange.getCell(j, 7).setValue(data[_buyer][_contract].overdueDebt);
      reportRange.getCell(j, 8).setValue(data[_buyer][_contract].debtDays);
      j++
    }
  }

} // func end

function generateInputs () {
  const buyers = new Set();
  const buyer_contracts = new Set();
  for (let i = 6; i < allSheets.length; i++ ) {
    const current_sheet = allSheets[i];
    const sheetName = current_sheet.getName();
    const idx = sheetName.indexOf('$');
    const companyName = idx === -1 ? sheetName : sheetName.slice(0, idx);
    
    buyers.add(companyName);
    buyer_contracts.add(sheetName);
  }
  

  const iterator_b = buyers.values();
  const iterator_c = buyer_contracts.values();
  

  const inputs = workbook.getSheetByName('Служебный')

  const input_buyers = inputs.getRange('A3:A');
  input_buyers.clearContent();
  
  for (let i = 1; i < buyers.size + 1 ; i++) {
    input_buyers.getCell(i, 1).setValue(iterator_b.next().value)
  }

  const input_contracts = inputs.getRange('B3:B');
  input_contracts.clearContent();

  for (let i = 1; i < buyers.size + 1 ; i++) {
    input_contracts.getCell(i, 1).setValue(iterator_c.next().value)
  }
}

function deleteUnnecessary () {
  const totalSheet = workbook.getSheetByName('ТОТАЛ');
  const nameRange = totalSheet.getRange('C7:C');
  const nameData = nameRange.getValues();
  const names = new Set();

  for (name of nameData) {
    if (name[0] !== '') {
      names.add(name[0]);
      // console.log('adding',name[0])
    }
  }
  
  for (sheet of allSheets) {
    const curName = sheet.getName();
    if (curName !== "ТОТАЛ" & curName !== "Служебный" & curName !== "Задолженность Отчет" & curName !== "General scripts" & curName !== "Wrong data") {
        if (!names.has(sheet.getName())) {
        // console.log('ready to delete',sheet.getName())
        workbook.deleteSheet(sheet);
      }
    }
    
  }
}
