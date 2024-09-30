function myFunction() {

  // Logger.log(getData("yunxi he"));

  // yunxi = new Armwrestler("yunxi he");
  // Logger.log(yunxi.rating1() + ',' + yunxi.deviation1() + ',' + yunxi.s);


  // alpha = new Armwrestler("alpha");
  
  // yunxi.update(alpha, 1);
  // yunxi.apply();

  // Logger.log(alpha.rating1() + ',' + alpha.deviation1() + ',' + alpha.s);

  // Logger.log(yunxi.rating1() + ',' + yunxi.deviation1() + ',' + yunxi.s);

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Armwrestling')
    .addItem('New Wrestler', 'newWrestler')
    .addItem('New Match', 'newMatch')
    .addItem('Resimulate', 'resimulate')
    .addToUi();
}

function newWrestler(name = "", sort = 1){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (name == "") {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Armwrestler Name");
    

    if (response.getSelectedButton() == ui.Button.OK && !response.getResponseText() == "") {
      name = response.getResponseText();
    }

    else {
      ui.alert("error");
      return;
    }
  }

  for (let i = 2; i != -1; i++) {
    if (sheet.getRange(i, 1).getValue() == name) {
      Logger.log("User found: " + name);
      break;
    }
    
    if (sheet.getRange(i, 1).isBlank()) {
      for (let j = i; j > 2; j--) {
        sheet.getRange(j, 1, 1, 4).setValues(sheet.getRange(j - 1, 1, 1, 4).getValues());
      }

      sheet.getRange(2, 1, 1, 4).setValues([[name, 1500, 350, 0.06]]);
      Logger.log("User added: " + name);
      break;
    }
  }


  if (sort) autoSortUsers();
  SpreadsheetApp.flush();
}

function getData(userName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  for (let i = 2; i != -1; i++) {
    if (sheet.getRange(i, 1).getValue() == userName) {
      Logger.log("User found: " + userName);
      return sheet.getRange(i, 1, 1, 4);
    }
    
    if (sheet.getRange(i, 1).isBlank()) {
      // for (let j = i; j > 2; j--) {
      //   sheet.getRange(j, 1, 1, 4).setValues(sheet.getRange(j - 1, 1, 1, 4).getValues());
      // }

      // sheet.getRange(2, 1, 1, 4).setValues([[userName, 1500, 350, 0.06]]);
      // Logger.log("User added: " + userName);
      // return sheet.getRange(2, 1, 1, 4);
      // Logger.log("User not found:" + userName);
      // ui.alert("User not found: " + userName);
      return sheet.getRange(i, 1);
    }
  }
}

class Armwrestler {
  constructor(name, rating = Armwrestler.kDefaultR, deviation = Armwrestler.kDefaultRD, volatility = Armwrestler.kDefaultS) {
    this.name = name;
    this.u = (rating - Armwrestler.kDefaultR) / Armwrestler.kScale;
    this.p = deviation / Armwrestler.kScale;
    this.s = volatility;
  }

  rating1(){
    return (this.u * Armwrestler.kScale) + Armwrestler.kDefaultR;
  }

  deviation1(){
    return this.p * Armwrestler.kScale;
  }

  G() {
    let scale = this.p / Math.PI;
    return 1.0 / Math.sqrt(1.0 + 3.0 * scale * scale);
  }

  E(g, rating) {
    let exponent = -1.0 * g * (rating.u - this.u);
    return 1.0 / (1.0 + Math.exp(exponent)); 
  }

  F(x, dS, pS, v, a, tS) {
    let eX = Math.exp(x);
    let num = eX * (dS - pS - v - eX);
    let den = pS + v + eX;

    return (num / (2.0 * den * den)) - ((x-a) / tS);
  }

  convergence(d, v, p, s) {
    let dS = d * d;
    let pS = p * p;
    let tS = Armwrestler.kSystemConst * Armwrestler.kSystemConst;
    
    let a = Math.log(s * s);
  
    let A = a;
    let B;
    let bTest = dS - pS - v;

    if (bTest > 0.0) B = Math.log(bTest);
    else {
      B = a - Armwrestler.kSystemConst;
      while (this.F(B, dS, pS, v, a, tS) < 0.0) B -= Armwrestler.kSystemConst;
    }

    let fA = this.F(A, dS, pS, v, a, tS);
    let fB = this.F(B, dS, pS, v, a, tS);
    while (Math.abs(B - A) > Armwrestler.kConvergence) {
      let C = A + (A - B) * fA / (fB - fA);
      let fC = this.F(C, dS, pS, v, a, tS);

      if (fC * fB < 0.0) {
          A = B;
          fA = fB;
      }
      else fA /= 2.0;

      B = C;
      fB = fC;
    }
    return A;
  }

  update(opponent, score) {
    let g = opponent.G();
    let e = opponent.E(g, this);
    
    let invV = g * g * e * (1.0 - e);
    let v = 1.0 / invV;

    let dInner = g * (score - e);
    let d = v * dInner;

    this.sPrime = Math.exp(this.convergence(d, v, this.p, this.s) / 2.0);
    this.pPrime = 1.0 / Math.sqrt((1.0 / (this.p * this.p + this.sPrime * this.sPrime)) + invV);
    this.uPrime = this.u + this.pPrime * this.pPrime * dInner;
  }

  apply() {
    this.u = this.uPrime;
    this.p = this.pPrime;
    this.s = this.sPrime;
  }

}

/// The default/initial rating value
Armwrestler.kDefaultR = 1500.0;

/// The default/initial deviation value
Armwrestler.kDefaultRD = 350.0;

/// The default/initial volatility value
Armwrestler.kDefaultS = 0.06;


/// The Glicko-1 to Glicko-2 scale factor
Armwrestler.kScale = 173.7178;

/// The system constant (tau)
Armwrestler.kSystemConst = 0.5;

/// The convergence constant (epsilon)
Armwrestler.kConvergence = 0.000001;

function newMatch(sort = 1){
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const response1 = ui.prompt("Winner");
  if (response1.getSelectedButton() != ui.Button.OK) {
    Logger.log('User canceled the prompt.');
    ui.alert("User canceled the prompt.");
    return;
  }

  if (getData(response1.getResponseText()).isBlank()) {
    ui.alert("User not found."); 
    return;
  }

  const response2 = ui.prompt("Loser");
  if (response2.getSelectedButton() != ui.Button.OK) {
    Logger.log('User canceled the prompt.');
    ui.alert("User canceled the prompt.");
    return;
  }

  if (getData(response2.getResponseText()).isBlank()) {
    ui.alert("User not found."); 
    return;
  }

  const response3 = ui.prompt("Score");
  if (response3.getResponseText() == "") {
    ui.alert("Please enter score."); 
    return;
  }

  let i = 3;
  for (; !sheet.getRange(i, 7).isBlank(); i++);
  for (let j = i; j > 3; j--) {
    sheet.getRange(j, 7, 1, 4).setValues(sheet.getRange(j - 1, 7, 1, 4).getValues());
  }

  // Set format to plain text before setting the value
  sheet.getRange(3, 10).setNumberFormat("@STRING@"); // Assuming score is being placed in column 10
  
  sheet.getRange(3, 7, 1, 4).setValues([
    [response1.getResponseText(), response2.getResponseText(), Utilities.formatDate(new Date(), "GMT+1", "MM/dd/YY"), response3.getResponseText()]
  ]);
  
  // Apply updates for users
  let data1 = getData(response1.getResponseText()).getValues();
  let data2 = getData(response2.getResponseText()).getValues();
  
  let user1 = new Armwrestler(data1[0][0], data1[0][1], data1[0][2], data1[0][3]);
  let user2 = new Armwrestler(data2[0][0], data2[0][1], data2[0][2], data2[0][3]);

  user1.update(user2, 1);
  user2.update(user1, 0);

  user1.apply();
  user2.apply();

  getData(user1.name).setValues([[user1.name, user1.rating1(), user1.deviation1(), user1.s]]);
  getData(user2.name).setValues([[user2.name, user2.rating1(), user2.deviation1(), user2.s]]);

  if (sort) autoSortUsers();
  SpreadsheetApp.flush();
}


function simulateMatch(name1, name2) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (getData(name1).isBlank()) newWrestler(name1, sort = 0);
  if (getData(name2).isBlank()) newWrestler(name2, sort = 0);


  let data1 = getData(name1).getValues();
  let data2 = getData(name2).getValues();
  
  let user1 = new Armwrestler(data1[0][0], data1[0][1], data1[0][2], data1[0][3]);
  let user2 = new Armwrestler(data2[0][0], data2[0][1], data2[0][2], data2[0][3]);


  user1.update(user2, 1);
  user2.update(user1, 0);

  user1.apply();
  user2.apply();

  // ui.alert(user1.name + " " + user2.name);
  // ui.alert(response1.getResponseText() + " " + response2.getResponseText());

  getData(user1.name).setValues([[user1.name, user1.rating1(), user1.deviation1(), user1.s]]);
  getData(user2.name).setValues([[user2.name, user2.rating1(), user2.deviation1(), user2.s]]);

}

function userCount() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  for (var i = 2; !sheet.getRange(i, 1).isBlank(); i++);
  Logger.log(i - 1);
  return i - 1;

}

function matchCount() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  for (var i = 3; !sheet.getRange(i, 7).isBlank(); i++);
  Logger.log("matchcount " + i - 1);
  return i - 1;
}

function autoSortUsers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let count = userCount();
  sheet.getRange(2, 1, count, 4).sort({column: 2, ascending: false});
  
}

function resimulate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let uCount = userCount();
  let mCount = matchCount();

  for (let i = 2; !sheet.getRange(i, 1, 1, 4).isBlank(); i++) {
    sheet.getRange(i, 1, 1, 4).clearContent();
  }

  for (let i = mCount; i > 2; i--) {
    row = sheet.getRange(i, 7, 1, 2);
    let name1 = row.getValues()[0][0];
    let name2 = row.getValues()[0][1];
    simulateMatch(name1, name2);
  }

  autoSortUsers();
  SpreadsheetApp.flush();
}


	
